@echo off
REM One-click setup for Windows: creates .venv, installs packages, creates .env
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 ( set "PY=py" ) else ( set "PY=python" )

%PY% --version || goto :nopython

if not exist .venv (
    echo Creating virtual environment...
    %PY% -m venv .venv || goto :fail
)
call .venv\Scripts\activate.bat || goto :fail

python -m pip install --upgrade pip || goto :fail
pip install -r requirements.txt || goto :fail
pip install -e . || goto :fail

if not exist .env (
    copy .env.example .env >nul
    echo Created .env - open it and fill in your credentials.
)
echo.
echo Setup complete. Next:
echo   1. Edit .env
echo   2. .venv\Scripts\activate.bat
echo   3. python scripts\test_db_connectivity.py
exit /b 0

:nopython
echo Python was not found. Install Python 3.11 or 3.12 from https://www.python.org/downloads/ and tick "Add python.exe to PATH".
exit /b 1
:fail
echo Setup failed - see the messages above.
exit /b 1
