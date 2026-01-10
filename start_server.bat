@echo off
cd /d %~dp0
echo Starting Flask server...
venv\Scripts\python.exe app.py
pause
