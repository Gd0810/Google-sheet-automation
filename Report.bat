@echo off

REM Go to project folder
cd /d C:\Users\redback\Downloads\Report

REM Activate virtual environment
call .venv\Scripts\activate

REM Start Flask app in background
start cmd /k python app.py

REM Wait for server to start
timeout /t 6 >nul

REM Open browser
start http://127.0.0.1:5000/

exit