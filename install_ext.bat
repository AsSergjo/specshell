@echo off
cd /d "%~dp0"
REM Registers the context menu handler.
REM This script must be run as an administrator.

REM The name of the DLL to register.
set DLL_NAME=SpectrogramContextMenu.dll

echo Registering %DLL_NAME%...
regsvr32 %DLL_NAME%

if %errorlevel% equ 0 (
    echo.
    echo Successfully registered the extension.
    echo.
    echo To apply the changes, Windows Explorer needs to be restarted.
    echo This will cause your desktop and taskbar to disappear for a moment.
    
) else (
    echo.
    echo Failed to register the extension. Make sure you are running this script as an Administrator.
	pause
	exit
)

choice /c YN /m "Do you want to restart Explorer now?"

if %errorlevel% equ 1 (
       echo Restarting Explorer...
        taskkill /f /im explorer.exe
        timeout /t 2 /nobreak >nul
        start explorer.exe
    )


echo Installation complete.

pause
