@echo off
echo Sideloading MacroFlow AI without debugging...

REM Get the project root directory (parent of scripts)
cd /d "%~dp0.."

REM Registry key for Office add-ins
set REGISTRY_KEY=HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\TrustedCatalogs\{fba090d8-b4fc-4f86-a421-fa32126eda1b}

REM Full path to manifest
set MANIFEST_PATH=%CD%\frontend\manifest.xml

echo Registering add-in manifest at: %MANIFEST_PATH%

REM Create trusted catalog entry (this avoids debugging mode)
reg add "%REGISTRY_KEY%" /v "Id" /t REG_SZ /d "{fba090d8-b4fc-4f86-a421-fa32126eda1b}" /f
reg add "%REGISTRY_KEY%" /v "Url" /t REG_SZ /d "file:///%MANIFEST_PATH%" /f
reg add "%REGISTRY_KEY%" /v "Flags" /t REG_DWORD /d 1 /f

REM Disable debug dialog for this add-in
echo Disabling debug dialog...
reg add "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\fba090d8-b4fc-4f86-a421-fa32126eda1a" /v "UseDirectDebugger" /t REG_DWORD /d 0 /f

echo.
echo MacroFlow AI has been sideloaded as a production add-in.
echo Please restart Excel to see the changes.
echo.
echo To remove the add-in, run: scripts\remove-sideload.bat
echo.
pause