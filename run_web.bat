@echo off
REM Launch the Outlook Assistant localhost dashboard.
REM Edit OUTLOOK_WEB_TOKEN below to require a token (recommended if not single-user).
cd /d "%~dp0"
set OUTLOOK_WEB_HOST=127.0.0.1
set OUTLOOK_WEB_PORT=8765
REM set OUTLOOK_WEB_TOKEN=change-me
python outlook_web.py
pause
