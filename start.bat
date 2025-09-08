@echo off
echo Starting MacroFlow AI...

REM Install frontend dependencies if needed
echo Checking frontend dependencies...
cd frontend
if not exist node_modules (
    echo Installing frontend dependencies...
    npm install
)

REM Start Flask backend in a new window
echo Starting Flask backend (port 5000)...
cd ..
start "MacroFlow Backend" cmd /k "py app.py"

REM Wait a moment for backend to start
timeout /t 3 /nobreak >nul

REM Start frontend server in a new window
echo Starting frontend server (port 3000)...
cd frontend
start "MacroFlow Frontend" cmd /k "node server.js"

REM Start Office add-in debugging
timeout /t 2 /nobreak >nul
start "MacroFlow Excel" cmd /k "npx office-addin-debugging start manifest.xml"

echo.
echo MacroFlow AI is starting up...
echo Backend: http://localhost:5000
echo Frontend: https://localhost:3000
echo.
echo All servers will open in separate command windows.
echo Close those windows to stop the servers.
echo.
pause