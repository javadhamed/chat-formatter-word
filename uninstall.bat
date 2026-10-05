@echo off
setlocal EnableExtensions
chcp 65001 >nul 2>&1

REM ==============================================================
REM   Chat Formatter for Word - Uninstaller (Windows)
REM ==============================================================

title Chat Formatter - Uninstall

set "STARTUP=%APPDATA%\Microsoft\Word\STARTUP"
set "CFHOME=%USERPROFILE%"
if not defined CFHOME set "CFHOME=%HOMEDRIVE%%HOMEPATH%"

echo.
echo ==============================================================
echo   Chat Formatter for Word - Uninstall
echo ==============================================================
echo.

echo [1/3] Checking for running Word...
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if not errorlevel 1 (
    echo       Word is running. Close it, then reopen Word.
) else (
    echo       Word is not running.
)

echo [2/3] Removing the template...
if exist "%STARTUP%\ChatFormatter.dotm" (
    del /F /Q "%STARTUP%\ChatFormatter.dotm" >nul 2>&1
    echo       Removed.
) else (
    echo       Nothing to remove.
)

echo [3/3] Removing your settings...
if exist "%CFHOME%\.chatformatter" (
    rmdir /S /Q "%CFHOME%\.chatformatter" >nul 2>&1
    echo       Removed.
) else (
    echo       No settings to remove.
)

echo.
echo The trusted location entry was left in place; it is harmless.
echo Word no longer has Chat Formatter loaded.
echo.
pause
endlocal
