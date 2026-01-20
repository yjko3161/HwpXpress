@echo off
if not exist ".venv" (
    echo [ERROR] Virtual environment not found. Please run setup_env.bat first.
    pause
    exit /b 1
)

echo [INFO] Activating virtual environment and starting HwpXpress...
call .venv\Scripts\activate
python main.py
