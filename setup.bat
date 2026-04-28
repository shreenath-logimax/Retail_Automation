@echo off
echo ===================================================
echo Setting up Retail Automation Project
echo ===================================================

echo.
echo [1/3] Creating virtual environment (.venv)...
python -m venv .venv
if %errorlevel% neq 0 (
    echo Error creating virtual environment. Please make sure Python is installed and in your PATH.
    pause
    exit /b %errorlevel%
)

echo.
echo [2/3] Activating virtual environment...
call .venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo Error activating virtual environment.
    pause
    exit /b %errorlevel%
)

echo.
echo [3/3] Installing dependencies from requirements.txt...
python -m pip install --upgrade pip
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo Error installing dependencies.
    pause
    exit /b %errorlevel%
)

echo.
echo ===================================================
echo Setup completed successfully!
echo.
echo To run the application, make sure to always activate
echo the virtual environment first by running:
echo     .venv\Scripts\activate
echo ===================================================
pause
