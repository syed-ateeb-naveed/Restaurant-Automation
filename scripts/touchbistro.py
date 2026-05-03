"""touchbistro.py
================
Consolidated TouchBistro reporting script.

Currently implemented
---------------------
  run_order_type()  —  Sales by Order Type

How it works
------------
1. Selenium starts a fresh Edge window and logs in to admin.touchbistro.com
   automatically using the credentials in the CONFIG section.
   No manual browser setup or debugger session required.
2. Makes API calls directly through the browser using execute_async_script,
   so the browser's own auth (Okta session, cookies, tokens) is used — no
   need to copy or replicate any of it.
3. Sums the daily records to get weekly totals.
4. Writes the results into the Excel sheet.
5. Closes the browser when done.

Usage
-----
1. Set PREV_WEEK_LABEL, TARGET_WEEK_LABEL, WEEK_START, and WEEK_END below.
2. Run:  python touchbistro.py
"""

from datetime import datetime, timezone, timedelta
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# ════════════════════════════════════════════════════════════════
# CONFIG  —  update these before each run
# ════════════════════════════════════════════════════════════════

DRIVER_PATH = "msedgedriver.exe"
MASTER_XLSX = r"..\Weekly Reports\Weekly Sales summary (week 5).xlsx"

# TouchBistro login credentials
TB_USERNAME = "muneef.naveed30@gmail.com"
TB_PASSWORD = "Bistro1#$"

# Label of the previous week — used to locate the correct position if the
# target row doesn't exist yet (the script will write into the row below it).
PREV_WEEK_LABEL   = "Feb 02 - Feb 08"

# Label for the week you are filling in.
# If this row already exists in column A, the script writes into it directly.
# If it doesn't exist, it creates the row below PREV_WEEK_LABEL.
TARGET_WEEK_LABEL = "Feb 09 - Feb 12"

# Date range for the report (YYYY-MM-DD), both ends inclusive.
# WEEK_START = first day of the week (e.g. Monday)
# WEEK_END   = last day of the week (e.g. Sunday)
# The script automatically adds 1 day to WEEK_END when calling the API,
# because the TouchBistro endpoint treats the end timestamp as exclusive.
WEEK_START = "2025-02-09"
WEEK_END   = "2025-02-12"


# ════════════════════════════════════════════════════════════════
# VENUE CONFIG
# ════════════════════════════════════════════════════════════════

LOGIN_URL = "https://login.touchbistro.com/login/login.htm"
ADMIN_URL = "https://admin.touchbistro.com"
ORDER_TYPE_URL = (
    ADMIN_URL
    + "/api/frontend/report/v1/venues/{venue_id}/reports/order-type"
    + "?start={start}&end={end}"
)

# venue_id → (display name, Excel sheet name)
VENUES = {
    56776: ("Karachi Kabab Wala - Queen Street", "Kababwala - Queen"),
    55119: ("Pizza Karachi- Eglinton",           "Pizza K Eglinton"),
    55118: ("Pizza Karachi -Heartland",          "Pizza K Heartland"),
    51879: ("Karachi Kabab Wala",                "Kababwala"),
    51876: ("Karachi Food Court",                "Karachi Food Court"),
    54708: ("Pizza Karachi Downtown TO",         "Queen St."),
    52043: ("Pizza Karachi- Highway Karahi",     "Highway"),
    51880: ("Pizza Karachi - Wonderland",        "Jane"),
    51878: ("Pizza Karachi - Lebovic",           "Lebovic"),
    51877: ("Pizza Karachi - Ajax",              "Ajax"),
    51594: ("Pizza Karachi - Markham Rd",        "Markham"),
}

# Excel column where "Take Out" lives (H = 8).
# Delivery / Bar Tab / Dine-In / Online Order / Net Sales follow in columns 9-13.
ORDER_TYPE_START_COL = 8


# ════════════════════════════════════════════════════════════════
# HELPERS
# ════════════════════════════════════════════════════════════════

def to_unix(date_str: str, add_day: bool = False) -> int:
    """
    Convert 'YYYY-MM-DD' to a Unix timestamp at midnight UTC.
    Pass add_day=True for the end date so the API's exclusive upper bound
    correctly includes the last day of the week.
    """
    dt = datetime.strptime(date_str, "%Y-%m-%d").replace(tzinfo=timezone.utc)
    if add_day:
        dt += timedelta(days=1)
    return int(dt.timestamp())


def login() -> webdriver.Edge:
    """
    Start a fresh Edge window, navigate to the TouchBistro login page,
    fill in credentials, and wait until we land on admin.touchbistro.com.
    Returns the authenticated driver.
    """
    opts = Options()
    opts.use_chromium = True
    driver = webdriver.Edge(service=Service(DRIVER_PATH), options=opts)
    wait = WebDriverWait(driver, 30)

    print("Opening login page...")
    driver.get(LOGIN_URL)

    # Fill in the Okta login form
    wait.until(EC.presence_of_element_located((By.ID, "okta-signin-username"))).send_keys(TB_USERNAME)
    driver.find_element(By.ID, "okta-signin-password").send_keys(TB_PASSWORD)
    driver.find_element(By.ID, "okta-signin-submit").click()

    # Wait until redirected to the admin panel
    print("Waiting for login to complete...")
    wait.until(EC.url_contains("admin.touchbistro.com"))
    print("✅ Logged in.\n")

    return driver


def fetch_via_browser(driver, url: str) -> list[dict]:
    """
    Make an API call through the browser using fetch(), so the browser's
    own Okta session and auth tokens are used automatically.
    Returns the parsed JSON response, or raises on error.
    """
    result = driver.execute_async_script("""
        var url      = arguments[0];
        var callback = arguments[1];
        fetch(url)
            .then(function(response) {
                if (!response.ok) {
                    callback({ __error__: response.status + ' ' + response.statusText });
                } else {
                    response.json().then(function(data) { callback(data); });
                }
            })
            .catch(function(err) {
                callback({ __error__: err.toString() });
            });
    """, url)

    if isinstance(result, dict) and "__error__" in result:
        raise RuntimeError(f"API error: {result['__error__']}")
    return result


def sum_totals(records: list[dict]) -> list[float]:
    """
    Sum daily records into a single weekly total.
    Returns: [Take Out, Delivery, Bar Tab, Dine-In, Online Order, Net Sales]
    """
    keys = ["takeout", "delivery", "bartab", "dinein", "onlineorder", "total"]
    return [round(sum(float(r[k]) for r in records), 2) for k in keys]


def find_target_row(ws) -> int:
    """
    Return the row number to write into, creating it if needed.

    1. If TARGET_WEEK_LABEL already exists in column A → use that row.
    2. Otherwise, find PREV_WEEK_LABEL and use the row directly below it,
       writing TARGET_WEEK_LABEL into column A of that row.
    3. If neither label is found → raise a clear error.
    """
    prev_row = None
    for cell in ws["A"]:
        val = str(cell.value).strip() if cell.value is not None else ""
        if val == TARGET_WEEK_LABEL:
            return cell.row          # already exists — nothing to create
        if val == PREV_WEEK_LABEL:
            prev_row = cell.row

    if prev_row is None:
        raise RuntimeError(
            f"Neither '{TARGET_WEEK_LABEL}' nor '{PREV_WEEK_LABEL}' found "
            f"in column A of sheet '{ws.title}'. Check the CONFIG labels."
        )

    # Target row doesn't exist yet — create it below the previous week
    target_row = prev_row + 1
    ws.cell(row=target_row, column=1, value=TARGET_WEEK_LABEL)
    print(f"   ℹ️  Row for '{TARGET_WEEK_LABEL}' not found — created at row {target_row}.")
    return target_row


# ════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════

def main():
    print(f"Week: {WEEK_START} → {WEEK_END}  |  Label: '{TARGET_WEEK_LABEL}'")

    driver = login()

    try:
        wb = load_workbook(MASTER_XLSX)

        start_ts = to_unix(WEEK_START)
        end_ts   = to_unix(WEEK_END, add_day=True)
        print(f"Timestamps: start={start_ts}  end={end_ts}\n")

        col_labels = ["Take Out", "Delivery", "Bar Tab", "Dine-In", "Online Order", "Net Sales"]

        for venue_id, (name, sheet_name) in VENUES.items():
            print(f"▶  {name}")

            url = ORDER_TYPE_URL.format(venue_id=venue_id, start=start_ts, end=end_ts)
            try:
                records = fetch_via_browser(driver, url)
            except RuntimeError as exc:
                print(f"   ⚠️  {exc} — skipping.")
                continue

            if not records:
                print("   ⚠️  No data returned for this date range — skipping.")
                continue

            values = sum_totals(records)
            for lbl, val in zip(col_labels, values):
                print(f"   {lbl:<14}: {val:>10.2f}")

            ws = wb[sheet_name]
            row = find_target_row(ws)
            for idx, val in enumerate(values):
                ws.cell(row=row, column=ORDER_TYPE_START_COL + idx, value=val)
            print(f"   ✅ Written to row {row}\n")

        wb.save(MASTER_XLSX)
        print(f"✅  Done. Workbook saved → {MASTER_XLSX}")

    finally:
        driver.quit()


if __name__ == "__main__":
    main()
