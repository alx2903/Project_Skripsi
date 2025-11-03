@echo off
title Flask Desktop Application
echo Starting the Flask application...

:: Set variables
set VENV_DIR=.venv
set VENV_PYTHON=%VENV_DIR%\Scripts\python.exe
set APP_SCRIPT=app.py
set SHORTCUT_NAME=Flask App.lnk
set ICON_PATH=dovechem_logo_iuP_icon.ico
set SHORTCUT_PATH=%USERPROFILE%\Desktop\%SHORTCUT_NAME%

:: Check if the virtual environment exists
if not exist "%VENV_PYTHON%" (
    echo Virtual environment not found. Creating a new one...
    python -m venv %VENV_DIR% --copies
    if %errorlevel% neq 0 (
        echo Failed to create virtual environment.
        pause
        exit /b
    )
    echo Virtual environment created successfully.
)

:: Activate the virtual environment
echo Activating virtual environment...
call %VENV_DIR%\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo Failed to activate virtual environment.
    pause
    exit /b
)

:: Install required packages
echo Installing required packages...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo Failed to install required packages.
    pause
    exit /b
)

:: Create a desktop shortcut if it doesn't already exist
if not exist "%SHORTCUT_PATH%" (
    echo Creating desktop shortcut...
    powershell -Command "$WshShell = New-Object -ComObject WScript.Shell; 
    $Shortcut = $WshShell.CreateShortcut('%SHORTCUT_PATH%'); 
    $Shortcut.TargetPath = '%CD%\%VENV_PYTHON%'; 
    $Shortcut.Arguments = '%CD%\%APP_SCRIPT%'; 
    $Shortcut.WorkingDirectory = '%CD%'; 
    $Shortcut.IconLocation = '%CD%\%ICON_PATH%'; 
    $Shortcut.Save()"
    echo Shortcut created on desktop.
) else (
    echo Shortcut already exists on desktop.
)

:: Run Flask application and wait until it starts
echo Starting Flask app...
start "Flask App" /D "%CD%" "%VENV_PYTHON%" "%CD%\%APP_SCRIPT%"

:: Wait for Flask to initialize
echo Waiting for Flask to start...
:WAIT_LOOP
timeout /t 3 >nul
netstat -ano | findstr :5000 >nul
if %errorlevel% neq 0 goto WAIT_LOOP

:: Open the browser after Flask starts
start http://127.0.0.1:5000

:: Exit
exit
