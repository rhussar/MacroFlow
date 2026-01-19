@echo off
echo Building MacroFlow for production...

REM Check if frontend dependencies are installed
echo Checking frontend dependencies...
cd frontend
if not exist node_modules (
    echo Installing frontend dependencies...
    npm install
)

REM Build frontend for production
echo Building frontend...
npm run build
if %ERRORLEVEL% neq 0 (
    echo Frontend build failed!
    pause
    exit /b 1
)

REM Install backend dependencies
echo Installing backend dependencies...
cd ..\backend
pip install -r requirements.txt
if %ERRORLEVEL% neq 0 (
    echo Backend dependency installation failed!
    pause
    exit /b 1
)

REM Validate Office manifest
echo Validating Office manifest...
cd ..\frontend
npm run validate
if %ERRORLEVEL% neq 0 (
    echo Manifest validation failed!
    pause
    exit /b 1
)

REM Create production configuration
echo Setting up production configuration...
cd ..
if not exist .env (
    echo Creating .env from template...
    copy .env.example .env
    echo Please edit .env file with your production settings!
    pause
)

echo.
echo MacroFlow build completed successfully!
echo.
echo Next steps:
echo 1. Configure your .env file with production settings
echo 2. Deploy backend/ folder to your Flask hosting service
echo 3. Deploy frontend/ folder to your web server with HTTPS
echo 4. Update manifest.xml URLs to point to your production domain
echo.
pause