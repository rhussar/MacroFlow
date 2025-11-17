@echo off
echo Removing MacroFlow sideload registrations...

REM Remove development registration and debug settings
echo Removing development registration and debug settings...
reg delete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\fba090d8-b4fc-4f86-a421-fa32126eda1a" /f 2>nul
reg delete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer" /v "fba090d8-b4fc-4f86-a421-fa32126eda1a" /f 2>nul

REM Remove trusted catalog registration
echo Removing trusted catalog registration...
reg delete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\TrustedCatalogs\{fba090d8-b4fc-4f86-a421-fa32126eda1b}" /f 2>nul

REM Remove runtime logging (if it exists)
echo Removing runtime logging...
powershell -Command "Remove-Item -Path 'HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\RuntimeLogging' -Recurse -Force -ErrorAction SilentlyContinue"

echo.
echo MacroFlow sideload registrations have been removed.
echo Please restart Excel to see the changes.
echo.
pause