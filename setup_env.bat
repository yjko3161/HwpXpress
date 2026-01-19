@echo off
setlocal

set VENV_DIR=.venv

echo [INFO] Checking for Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in your PATH.
    echo Please install Python and try again.
    pause
    exit /b 1
)

if not exist "%VENV_DIR%" (
    echo [INFO] Creating virtual environment...
    python -m venv %VENV_DIR%
) else (
    echo [INFO] Virtual environment already exists.
)

echo [INFO] Activating virtual environment...
call %VENV_DIR%\Scripts\activate

echo [INFO] Upgrading pip...
python -m pip install --upgrade pip

echo [INFO] Installing dependencies...
if exist requirements.txt (
    pip install -r requirements.txt
) else (
    echo [WARNING] requirements.txt not found.
)

echo.
echo [INFO] Setup complete. You can now run the application with 'python main.py'.
echo.
pause
