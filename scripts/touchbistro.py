"""touchbistro.py
================
Consolidated TouchBistro reporting script.

Reports implemented
-------------------
  run_order_type()       —  Sales by Order Type  (cols H–M)
  run_sales_by_section() —  Sales by Section     (cols B–G)
  run_sales_by_category()—  Sales by Category    (cols N+, dynamic)
  run_sales_by_hour()    —  Sales by Hour        (cols under "Daily and Hourly Sales" header)

How it works
------------
1. Selenium starts a fresh Edge window and logs in automatically.
2. Every API call is made through the browser via execute_async_script so
   the browser's own Okta session handles auth — nothing to copy or manage.
3. Results are summed from daily records (the API has no summary row).
4. Data is written into the correct Excel columns and the workbook is saved.

Usage
-----
1. Set PREV_WEEK_LABEL, TARGET_WEEK_LABEL, WEEK_START, WEEK_END below.
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

# Label of the previous week — used if the target row doesn't exist yet.
PREV_WEEK_LABEL   = "Apr 27 - May 03"

# Label for the week being filled in.
# If this row already exists in column A the script writes into it directly.
# If it doesn't exist, it creates the row below PREV_WEEK_LABEL.
TARGET_WEEK_LABEL = "Apr 27 - May 03"

# Date range (YYYY-MM-DD), both ends inclusive.
# WEEK_START = first day of the week, WEEK_END = last day of the week.
# The script adds 1 day to WEEK_END when calling the API (exclusive upper bound).
WEEK_START = "2026-04-27"
WEEK_END   = "2026-05-03"


# ════════════════════════════════════════════════════════════════
# VENUE CONFIG
# ════════════════════════════════════════════════════════════════

LOGIN_URL = "https://login.touchbistro.com/login/login.htm"
ADMIN_URL = "https://admin.touchbistro.com"

# API URL templates  —  {venue_id}, {start}, {end} are filled at runtime
ORDER_TYPE_URL = (
    ADMIN_URL + "/api/frontend/report/v1/venues/{venue_id}/reports/order-type"
    + "?start={start}&end={end}"
)
SECTION_URL = (
    ADMIN_URL + "/api/frontend/report/v1/venues/{venue_id}/reports/sales-by-section"
    + "?start={start}&end={end}"
)
CATEGORY_URL = (
    ADMIN_URL + "/api/frontend/report/v1/venues/{venue_id}/reports/sales-by-sales-category"
    + "?start={start}&end={end}&orderType=any"
)
HOURLY_URL = (
    ADMIN_URL + "/api/frontend/report/v1/venues/{venue_id}/reports/sales-hourly-net-bill-start"
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

# Time buckets for the hourly report (label, [hour strings from API])
HOUR_BUCKETS = [
    ("8AM - 11AM",  ["8am - 9am", "9am - 10am", "10am - 11am"]),
    ("11AM - 3PM",  ["11am - 12pm", "12pm - 1pm", "1pm - 2pm", "2pm - 3pm"]),
    ("3PM - 6PM",   ["3pm - 4pm", "4pm - 5pm", "5pm - 6pm"]),
    ("6PM - 11PM",  ["6pm - 7pm", "7pm - 8pm", "8pm - 9pm", "9pm - 10pm", "10pm - 11pm"]),
    ("11PM - 7AM",  ["11pm - 12am", "12am - 1am", "1am - 2am", "2am - 3am",
                     "3am - 4am", "4am - 5am", "5am - 6am", "6am - 7am"]),
]

# Delivery platform section names as they appear in the API response
DELIVERY_PLATFORMS = {
    "skip the dishes": "skip",
    "doordash":        "dd",
    "ubereats":        "uber",
}

# Excel column constants (1-based)
SECTION_BILL_COUNT_COL  = 2   # B
SECTION_NET_SALES_COL   = 3   # C
SECTION_REMAINING_COL   = 4   # D
SECTION_SKIP_COL        = 5   # E
SECTION_DD_COL          = 6   # F
SECTION_UBER_COL        = 7   # G
ORDER_TYPE_START_COL    = 8   # H  (Take Out … Net Sales → H–M)


# ════════════════════════════════════════════════════════════════
# HELPERS — general
# ════════════════════════════════════════════════════════════════

def to_unix(date_str: str, add_day: bool = False) -> int:
    """Convert 'YYYY-MM-DD' to a Unix timestamp at midnight UTC."""
    dt = datetime.strptime(date_str, "%Y-%m-%d").replace(tzinfo=timezone.utc)
    if add_day:
        dt += timedelta(days=1)
    return int(dt.timestamp())


def login() -> webdriver.Edge:
    """Start a fresh Edge window and log in to admin.touchbistro.com."""
    opts = Options()
    opts.use_chromium = True
    driver = webdriver.Edge(service=Service(DRIVER_PATH), options=opts)
    wait = WebDriverWait(driver, 30)

    print("Opening login page...")
    driver.get(LOGIN_URL)

    wait.until(EC.presence_of_element_located((By.ID, "okta-signin-username"))).send_keys(TB_USERNAME)
    driver.find_element(By.ID, "okta-signin-password").send_keys(TB_PASSWORD)
    driver.find_element(By.ID, "okta-signin-submit").click()

    print("Waiting for login to complete...")
    wait.until(EC.url_contains("admin.touchbistro.com"))
    print("✅ Logged in.\n")
    return driver


def fetch_via_browser(driver, url: str) -> list[dict]:
    """
    Make a GET request through the browser so its Okta session is used.
    Returns parsed JSON or raises RuntimeError on HTTP/network error.
    """
    result = driver.execute_async_script("""
        var url      = arguments[0];
        var callback = arguments[1];
        fetch(url)
            .then(function(r) {
                if (!r.ok) {
                    callback({ __error__: r.status + ' ' + r.statusText });
                } else {
                    r.json().then(function(d) { callback(d); });
                }
            })
            .catch(function(e) { callback({ __error__: e.toString() }); });
    """, url)

    if isinstance(result, dict) and "__error__" in result:
        raise RuntimeError(f"API error: {result['__error__']}")
    return result


def find_target_row(ws) -> int:
    """
    Return the row to write into.
    Uses TARGET_WEEK_LABEL if it exists; otherwise creates a row below PREV_WEEK_LABEL.
    """
    prev_row = None
    for cell in ws["A"]:
        val = str(cell.value).strip() if cell.value is not None else ""
        if val == TARGET_WEEK_LABEL:
            return cell.row
        if val == PREV_WEEK_LABEL:
            prev_row = cell.row

    if prev_row is None:
        raise RuntimeError(
            f"Neither '{TARGET_WEEK_LABEL}' nor '{PREV_WEEK_LABEL}' found "
            f"in column A of sheet '{ws.title}'. Check CONFIG labels."
        )
    target_row = prev_row + 1
    ws.cell(row=target_row, column=1, value=TARGET_WEEK_LABEL)
    print(f"   ℹ️  Created row {target_row} for '{TARGET_WEEK_LABEL}'.")
    return target_row


def build_url(template: str, venue_id: int, start_ts: int, end_ts: int) -> str:
    return template.format(venue_id=venue_id, start=start_ts, end=end_ts)


# ════════════════════════════════════════════════════════════════
# REPORT 1 — Sales by Order Type
# ════════════════════════════════════════════════════════════════

def run_order_type(driver, wb, start_ts, end_ts):
    """
    Sums daily order-type records and writes to cols H–M:
      Take Out | Delivery | Bar Tab | Dine-In | Online Order | Net Sales
    """
    print("── Sales by Order Type ──────────────────────────")
    keys = ["takeout", "delivery", "bartab", "dinein", "onlineorder", "total"]
    col_labels = ["Take Out", "Delivery", "Bar Tab", "Dine-In", "Online Order", "Net Sales"]

    for venue_id, (name, sheet_name) in VENUES.items():
        print(f"▶  {name}")
        url = build_url(ORDER_TYPE_URL, venue_id, start_ts, end_ts)
        try:
            records = fetch_via_browser(driver, url)
        except RuntimeError as exc:
            print(f"   ⚠️  {exc} — skipping."); continue

        if not records:
            print("   ⚠️  No data — skipping."); continue

        values = [round(sum(float(r[k]) for r in records), 2) for k in keys]
        for lbl, val in zip(col_labels, values):
            print(f"   {lbl:<14}: {val:>10.2f}")

        ws = wb[sheet_name]
        row = find_target_row(ws)
        for idx, val in enumerate(values):
            ws.cell(row=row, column=ORDER_TYPE_START_COL + idx, value=val)
        print(f"   ✅ Written to row {row}")


# ════════════════════════════════════════════════════════════════
# REPORT 2 — Sales by Section
# ════════════════════════════════════════════════════════════════

def run_sales_by_section(driver, wb, start_ts, end_ts):
    """
    Classifies sections into delivery platforms vs. in-house and writes to cols B–G:
      Total Bill Count | Net Sales | Remaining | Skip | DoorDash | Uber
    """
    print("\n── Sales by Section ─────────────────────────────")

    for venue_id, (name, sheet_name) in VENUES.items():
        print(f"▶  {name}")
        url = build_url(SECTION_URL, venue_id, start_ts, end_ts)
        try:
            records = fetch_via_browser(driver, url)
        except RuntimeError as exc:
            print(f"   ⚠️  {exc} — skipping."); continue

        if not records:
            print("   ⚠️  No data — skipping."); continue

        total_bill_count = sum(r["bill_count"] for r in records)
        net_sales = skip = dd = uber = remaining = 0.0

        for r in records:
            rev  = float(r["sales_revenue"])
            name_lower = r["section_name"].strip().lower()
            net_sales += rev
            if name_lower == "skip the dishes":
                skip += rev
            elif name_lower == "doordash":
                dd += rev
            elif name_lower == "ubereats":
                uber += rev
            else:
                remaining += rev

        print(f"   {'Bill Count':<14}: {total_bill_count:>10}")
        print(f"   {'Net Sales':<14}: {net_sales:>10.2f}")
        print(f"   {'Remaining':<14}: {remaining:>10.2f}")
        print(f"   {'Skip':<14}: {skip:>10.2f}")
        print(f"   {'DoorDash':<14}: {dd:>10.2f}")
        print(f"   {'Uber':<14}: {uber:>10.2f}")

        ws = wb[sheet_name]
        row = find_target_row(ws)
        ws.cell(row=row, column=SECTION_BILL_COUNT_COL, value=total_bill_count)
        ws.cell(row=row, column=SECTION_NET_SALES_COL,  value=round(net_sales, 2))
        ws.cell(row=row, column=SECTION_REMAINING_COL,  value=round(remaining, 2))
        ws.cell(row=row, column=SECTION_SKIP_COL,       value=round(skip, 2))
        ws.cell(row=row, column=SECTION_DD_COL,         value=round(dd, 2))
        ws.cell(row=row, column=SECTION_UBER_COL,       value=round(uber, 2))
        print(f"   ✅ Written to row {row}")


# ════════════════════════════════════════════════════════════════
# REPORT 3 — Sales by Category
# ════════════════════════════════════════════════════════════════

def _build_category_header_map(ws) -> dict[str, int]:
    """
    Scan row 2 of the sheet for column headers.
    Returns {header_lower: column_index} for every non-empty header cell.
    """
    header_map = {}
    for cell in ws[2]:
        if cell.value is not None:
            header_map[str(cell.value).strip().lower()] = cell.column
    return header_map


def _match_category(category_name: str, header_map: dict[str, int]) -> int | None:
    """
    Find the column for a category by checking if any Excel header
    starts with the category name (case-insensitive), e.g. header
    "Food (Pizza K)" matches API category "Food".
    """
    cat_lower = category_name.strip().lower()
    for header, col in header_map.items():
        if header.startswith(cat_lower):
            return col
    return None


def run_sales_by_category(driver, wb, start_ts, end_ts):
    """
    Writes each category's net_revenue to its matching Excel column.
    Categories present in the API but absent from the sheet are skipped
    with a warning (mirrors the original script's behaviour).
    """
    print("\n── Sales by Category ────────────────────────────")

    for venue_id, (name, sheet_name) in VENUES.items():
        print(f"▶  {name}")
        url = build_url(CATEGORY_URL, venue_id, start_ts, end_ts)
        try:
            records = fetch_via_browser(driver, url)
        except RuntimeError as exc:
            print(f"   ⚠️  {exc} — skipping."); continue

        if not records:
            print("   ⚠️  No data — skipping."); continue

        ws = wb[sheet_name]
        row = find_target_row(ws)
        header_map = _build_category_header_map(ws)

        for r in records:
            cat_name  = r["sales_category_name"]
            net_rev   = round(float(r["net_revenue"]), 2)
            col = _match_category(cat_name, header_map)

            if col is None:
                print(f"   ⚠️  Category '{cat_name}' not found in sheet headers — skipping.")
                continue

            ws.cell(row=row, column=col, value=net_rev)
            print(f"   {cat_name:<20}: {net_rev:>10.2f}  → col {col}")

        print(f"   ✅ Written to row {row}")


# ════════════════════════════════════════════════════════════════
# REPORT 4 — Sales by Hour
# ════════════════════════════════════════════════════════════════

_DAY_KEYS = [
    "sunday_net_sales", "monday_net_sales", "tuesday_net_sales",
    "wednesday_net_sales", "thursday_net_sales", "friday_net_sales",
    "saturday_net_sales",
]


def _find_hourly_start_col(ws) -> int:
    """
    Find the starting column of the 'Daily and Hourly Sales (TBD)'
    merged header in row 1. Raises if not found.
    """
    for m in ws.merged_cells.ranges:
        if m.min_row == 1:
            cell_val = ws.cell(m.min_row, m.min_col).value
            if cell_val and "hourly" in str(cell_val).lower():
                return m.min_col
    raise RuntimeError(
        f"'Daily and Hourly Sales' header not found in row 1 of sheet '{ws.title}'."
    )


def run_sales_by_hour(driver, wb, start_ts, end_ts):
    """
    Sums all 7 day columns for each hour slot, buckets them into 5 time
    ranges, and writes to the columns under the 'Daily and Hourly Sales'
    merged header.
    """
    print("\n── Sales by Hour ────────────────────────────────")

    for venue_id, (name, sheet_name) in VENUES.items():
        print(f"▶  {name}")
        url = build_url(HOURLY_URL, venue_id, start_ts, end_ts)
        try:
            records = fetch_via_browser(driver, url)
        except RuntimeError as exc:
            print(f"   ⚠️  {exc} — skipping."); continue

        if not records:
            print("   ⚠️  No data — skipping."); continue

        # Sum all 7 day columns per hour to get the week total for that slot
        hour_sums: dict[str, float] = {}
        for r in records:
            hour = r["hour_of_the_day"].strip().lower()
            hour_sums[hour] = sum(float(r[d]) for d in _DAY_KEYS)

        # Aggregate into time buckets
        bucket_values = []
        for bucket_label, hours in HOUR_BUCKETS:
            total = 0.0
            for h in hours:
                if h not in hour_sums:
                    print(f"   ℹ️  Hour '{h}' missing for {sheet_name}, treating as 0.")
                total += hour_sums.get(h, 0.0)
            bucket_values.append(round(total, 2))
            print(f"   {bucket_label:<14}: {bucket_values[-1]:>10.2f}")

        ws = wb[sheet_name]
        row = find_target_row(ws)

        try:
            start_col = _find_hourly_start_col(ws)
        except RuntimeError as exc:
            print(f"   ⚠️  {exc} — skipping."); continue

        for idx, val in enumerate(bucket_values):
            ws.cell(row=row, column=start_col + idx, value=val)
        print(f"   ✅ Written to row {row} starting at col {start_col}")


# ════════════════════════════════════════════════════════════════
# MAIN
# ════════════════════════════════════════════════════════════════

def main():
    print(f"Week : {WEEK_START} → {WEEK_END}")
    print(f"Label: '{TARGET_WEEK_LABEL}'\n")

    start_ts = to_unix(WEEK_START)
    end_ts   = to_unix(WEEK_END, add_day=True)
    print(f"Timestamps: start={start_ts}  end={end_ts}\n")

    driver = login()

    try:
        wb = load_workbook(MASTER_XLSX)

        run_order_type(driver, wb, start_ts, end_ts)
        run_sales_by_section(driver, wb, start_ts, end_ts)
        run_sales_by_category(driver, wb, start_ts, end_ts)
        run_sales_by_hour(driver, wb, start_ts, end_ts)

        wb.save(MASTER_XLSX)
        print(f"\n✅  All reports done. Workbook saved → {MASTER_XLSX}")

    finally:
        driver.quit()


if __name__ == "__main__":
    main()
