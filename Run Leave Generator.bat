@echo off
REM Double-click this file to launch the Leave Document Generator app.
cd /d "%~dp0"
pythonw leave_generator_app.py
if errorlevel 1 (
    echo.
    echo Something went wrong starting the app. Trying with visible console...
    python leave_generator_app.py
    pause
)
