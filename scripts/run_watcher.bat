@echo off
REM ──────────────────────────────────────────────────────
REM  run_watcher.bat  —  Starts the file watcher at login
REM  Runs hidden via Task Scheduler (At Log On trigger).
REM ──────────────────────────────────────────────────────

cd /d "%~dp0"

python watcher.py
