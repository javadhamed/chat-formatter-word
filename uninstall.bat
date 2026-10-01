@echo off
setlocal EnableExtensions EnableDelayedExpansion

REM ==============================================================
REM   Chat Formatter for Word - Uninstaller
REM   Removes the template, the trusted location and the settings.
REM ==============================================================

title Chat Formatter - Uninstall

set "TEMPLATE_NAME=ChatFormatter.dotm"
set "STARTUP=%APPDATA%\Microsoft\Word\STARTUP"
set "SECKEY=HKCU\Software\Microsoft\Office\16.0\Word\Security\Trusted Locations"

echo [1/3] Removing template...
if exist "%STARTUP%\%TEMPLATE_NAME%" (
    del /F /Q "%STARTUP%\%TEMPLATE_NAME%" >nul 2>&1
    if errorlevel 1 (
        echo [WARN] Could not delete the file. Close Word and try again.
    ) else (
        echo       OK
    )
) else (
    echo       Not installed.
)

echo [2/3] Removing trusted location...
reg delete "%SECKEY%\LocationChatFormatter" /f >nul 2>&1
if errorlevel 1 (
    echo       Not present.
) else (
    echo       OK
)

echo [3/3] Removing settings...
reg delete "HKCU\Software\ChatFormatter" /f >nul 2>&1
if errorlevel 1 (
    echo       Not present.
) else (
    echo       OK
)

echo.
echo Uninstalled. Restart Word to finish.
echo.
echo Note: the toolbar button created by the add-in disappears after
echo       Word restarts.
echo.
pause
endlocal