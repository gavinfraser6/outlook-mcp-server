@echo off
REM Generate the inbox digest (read-only). Point Windows Task Scheduler at this.
REM Add flags as desired, e.g.:  --auto-categorize   or   --email-to you@example.com
cd /d "%~dp0"
python outlook_schedule.py --days 2 --top 15 %*
