@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul 2>&1

REM ==============================================================
REM   Chat Formatter for Word - Installer
REM   Copies the template into Word's STARTUP folder and registers
REM   it as a trusted location so macros run without prompts.
REM ==============================================================

title Chat Formatter - Install

set "TEMPLATE_NAME=ChatFormatter.dotm"
set "SCRIPT_DIR=%~dp0"
set "SOURCE=%SCRIPT_DIR%%TEMPLATE_NAME%"

if not exist "%SOURCE%" (
    echo [ERROR] "%TEMPLATE_NAME%" not found next to this script.
    echo         Expected: %SOURCE%
    echo.
    echo         Download the full release ZIP, not just this file.
    pause
    exit /b 1
)

REM --- locate Word STARTUP folder --------------------------------
set "STARTUP=%APPDATA%\Microsoft\Word\STARTUP"
if not exist "%STARTUP%" (
    echo [1/4] Creating STARTUP folder: %STARTUP%
    mkdir "%STARTUP%" 2>nul
)

echo [1/4] Copying template to STARTUP folder...
copy /Y "%SOURCE%" "%STARTUP%\%TEMPLATE_NAME%" >nul
if errorlevel 1 (
    echo [ERROR] Copy failed. Try running this file as Administrator.
    pause
    exit /b 1
)
echo       OK -^> %STARTUP%\%TEMPLATE_NAME%

REM --- register STARTUP as a trusted location ---------------------
REM Without this, Word blocks macros from the add-in under the
REM default "Disable all macros with notification" setting.
echo [2/4] Registering trusted location...
set "SECKEY=HKCU\Software\Microsoft\Office\16.0\Word\Security\Trusted Locations"
reg query "%SECKEY%" >nul 2>&1
if errorlevel 1 (
    reg add "%SECKEY%" /f >nul 2>&1
)
reg add "%SECKEY%\LocationChatFormatter" /v Path /t REG_EXPAND_SZ /d "%STARTUP%" /f >nul
reg add "%SECKEY%\LocationChatFormatter" /v Description /t REG_SZ /d "Chat Formatter for Word" /f >nul
reg add "%SECKEY%\LocationChatFormatter" /v AllowSubfolders /t REG_DWORD /d 1 /f >nul
if errorlevel 1 (
    echo [WARN] Could not register the trusted location automatically.
    echo        Add this folder manually in Word:
    echo        File ^> Options ^> Trust Center ^> Trust Center Settings
    echo        ^> Trusted Locations ^> Add new location
) else (
    echo       OK
)

REM --- default settings -------------------------------------------
echo [3/4] Writing default settings...
reg add "HKCU\Software\ChatFormatter" /v AutoFormat /t REG_SZ /d "1" /f >nul
reg add "HKCU\Software\ChatFormatter" /v Tables /t REG_SZ /d "1" /f >nul
echo       OK

REM --- close Word so the change takes effect ------------------------
echo [4/4] Checking for running Word...
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if not errorlevel 1 (
    echo       Word is running. Close it, then reopen Word.
) else (
    echo       Word is not running.
)

echo.
echo ==============================================================
echo   Installation complete.
echo ==============================================================
echo.
echo   How to use:
echo     1. Copy text from an AI chat (DeepSeek, ChatGPT, ...)
echo     2. Paste it into Word
echo     3. Select the pasted text and RIGHT-CLICK
echo.
echo   It formats automatically: headings, bold, tables, lists, code.
echo.
echo   Prefer an explicit button? Alt+F8 -> FormatChatText
echo   Want to turn the automatic mode off? Alt+F8 -^> ToggleAutoFormat
echo.
pause
endlocal