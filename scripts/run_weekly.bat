@echo off
REM ──────────────────────────────────────────────────────
REM  run_weekly.bat  —  Weekly TouchBistro report launcher
REM  Designed to be called by Windows Task Scheduler.
REM ──────────────────────────────────────────────────────

cd /d "%~dp0"

echo [%date% %time%] Starting TouchBistro weekly report...
python touchbistro.py
set EXIT_CODE=%ERRORLEVEL%

if %EXIT_CODE% NEQ 0 (
    echo [%date% %time%] ERROR: touchbistro.py exited with code %EXIT_CODE%
) else (
    echo [%date% %time%] SUCCESS: Weekly report completed.
)

exit /b %EXIT_CODE%
