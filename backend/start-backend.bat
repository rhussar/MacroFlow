@echo off
echo ========================================
echo   MacroFlow Backend Server
echo ========================================
echo.
echo Starting Flask server on http://localhost:5000
echo.
echo Press Ctrl+C to stop the server
echo.

cd /d "%~dp0"
py app.py

pause
