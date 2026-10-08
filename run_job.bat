@echo off
REM Wrapper for Task Scheduler / double-click use:   run_job.bat <job> [options]
cd /d "%~dp0"
call .venv\Scripts\activate.bat
python scripts\run_job.py %*
exit /b %ERRORLEVEL%
