@echo off
setlocal EnableExtensions

REM ==============================================================
REM   Chat Formatter for Word - Installer (Windows)
REM   Automatically removes any previous version first, then
REM   copies ChatFormatter.dotm into Word's STARTUP folder,
REM   registers Trusted Locations, and writes default settings.
REM ==============================================================

title Chat Formatter - Install

set "TEMPLATE_NAME=ChatFormatter.dotm"
set "SCRIPT_DIR=%~dp0"
set "SOURCE=%SCRIPT_DIR%%TEMPLATE_NAME%"
set "CF_EXIT=0"

echo.
echo   User: %USERNAME%

if not exist "%SOURCE%" (
    echo [ERROR] "%TEMPLATE_NAME%" not found next to this script.
    echo         Expected: %SOURCE%
    echo.
    echo         Use ChatFormatter-Windows-Installer.bat for standalone
    echo         single-file install, or extract the full release ZIP.
    set "CF_EXIT=1"
    goto :Cleanup
)

set "CF_SRC_SIZE="
for %%F in ("%SOURCE%") do set "CF_SRC_SIZE=%%~zF"

REM --- 1. ensure Word is closed so STARTUP isn't locked ----------
echo [1/4] Checking if Microsoft Word is running...
call :EnsureWordClosed
if errorlevel 1 (
    set "CF_EXIT=1"
    goto :Cleanup
)

REM --- 2. locate Word STARTUP and remove any previous version ----
echo [2/4] Removing any previous version...
set "DEFAULT_STARTUP=%APPDATA%\Microsoft\Word\STARTUP"
set "TARGET_STARTUP=%DEFAULT_STARTUP%"
set "ACTIVE_VER="
set "REG_STARTUP="

call :CleanDir "%DEFAULT_STARTUP%"
for %%V in (16.0 15.0 14.0) do (
    reg query "HKCU\Software\Microsoft\Office\%%V\Word" >nul 2>&1
    if not errorlevel 1 (
        call :CheckVersionStartup "%%V"
    )
)

set "CFHOME=%USERPROFILE%"
if not defined CFHOME set "CFHOME=%HOMEDRIVE%%HOMEPATH%"

call :CleanDir "%LOCALAPPDATA%\Microsoft\Word\STARTUP"
call :CleanDir "%LOCALAPPDATA%\Microsoft\Office\Word\STARTUP"
call :CleanDir "%CFHOME%\Documents\Custom Office Templates"

reg delete "HKCU\Software\ChatFormatter" /f >nul 2>&1
for %%V in (14.0 15.0 16.0) do (
    reg delete "HKCU\Software\Microsoft\Office\%%V\Word\Security\Trusted Locations\LocationChatFormatter" /f >nul 2>&1
)

if exist "%CFHOME%\.chatformatter" (
    rmdir /S /Q "%CFHOME%\.chatformatter" >nul 2>&1
)
echo       Previous files and settings cleaned.

REM --- 3. copy new template into STARTUP -------------------------
echo [3/4] Installing new template to STARTUP folder...
echo       Target folder: %TARGET_STARTUP%
if not exist "%TARGET_STARTUP%" mkdir "%TARGET_STARTUP%" 2>nul
if not exist "%TARGET_STARTUP%" (
    echo       [WARN] Could not access %TARGET_STARTUP%, falling back to default STARTUP.
    set "TARGET_STARTUP=%DEFAULT_STARTUP%"
    if not exist "%TARGET_STARTUP%" mkdir "%TARGET_STARTUP%" 2>nul
)
if not exist "%TARGET_STARTUP%" (
    echo [ERROR] Could not create STARTUP folder: %TARGET_STARTUP%
    set "CF_EXIT=1"
    goto :Cleanup
)
copy /Y "%SOURCE%" "%TARGET_STARTUP%\%TEMPLATE_NAME%" >nul
if errorlevel 1 (
    echo [ERROR] Copy failed. Close Word and try again.
    set "CF_EXIT=1"
    goto :Cleanup
)
set "CF_INST_SIZE="
for %%F in ("%TARGET_STARTUP%\%TEMPLATE_NAME%") do set "CF_INST_SIZE=%%~zF"
if not "%CF_INST_SIZE%"=="%CF_SRC_SIZE%" (
    echo [ERROR] Installed file verification failed at:
    echo         %TARGET_STARTUP%\%TEMPLATE_NAME%
    set "CF_EXIT=1"
    goto :Cleanup
)
echo       OK -^> %TARGET_STARTUP%\%TEMPLATE_NAME% ^(%CF_INST_SIZE% bytes^)

REM --- 4. register STARTUP as a trusted location & write settings
echo [4/4] Registering trusted location and default settings...
set "REG_OK=0"
for %%V in (14.0 15.0 16.0) do (
    reg query "HKCU\Software\Microsoft\Office\%%V\Word" >nul 2>&1
    if not errorlevel 1 (
        call :RegisterTrust "%%V" "%TARGET_STARTUP%"
    )
)
if "%REG_OK%"=="0" call :RegisterTrust "16.0" "%TARGET_STARTUP%"
if "%REG_OK%"=="1" (
    echo       Trusted location registered.
) else (
    echo       [WARN] Could not register the location automatically.
    echo        Add it by hand in Word:
    echo        File ^> Options ^> Trust Center ^> Trust Center Settings
    echo        ^> Trusted Locations ^> Add "%TARGET_STARTUP%"
)

if not exist "%CFHOME%\.chatformatter" mkdir "%CFHOME%\.chatformatter" 2>nul
(
    echo Tables = 1
) > "%CFHOME%\.chatformatter\settings.ini"
if exist "%CFHOME%\.chatformatter\settings.ini" (
    echo       Created %CFHOME%\.chatformatter\settings.ini
)

echo.
echo ==============================================================
echo   Installation complete.
echo ==============================================================
echo.
echo   Installed: %TARGET_STARTUP%\%TEMPLATE_NAME%
echo.
echo   How to use:
echo     1. Open Word
echo     2. Copy text from an AI chat ^(ChatGPT, DeepSeek, ...^)
echo     3. Paste it into Word
echo     4. Open the Add-ins tab and click "Format Chat Text"
echo.
echo   Other macros ^(Alt+F8^): FormatSelection, ShowSettings
echo.
goto :Cleanup

:EnsureWordClosed
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if errorlevel 1 (
    echo       Word is not running.
    exit /b 0
)
echo       Word is running. Requesting graceful close...
taskkill /IM WINWORD.EXE >nul 2>&1
for /L %%I in (1,1,10) do (
    ping -n 2 127.0.0.1 >nul
    tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
    if errorlevel 1 (
        echo       Word closed.
        exit /b 0
    )
)
echo       Word is still open ^(you may have an unsaved document^).
:WordPrompt
set "CF_ANS="
set /p "CF_ANS=      Save/close Word then press Enter to retry, [F] to force-close, or [Q] to quit: "
if /I "%CF_ANS%"=="Q" exit /b 1
if /I "%CF_ANS%"=="F" (
    taskkill /F /IM WINWORD.EXE >nul 2>&1
    ping -n 2 127.0.0.1 >nul
)
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if not errorlevel 1 goto :WordPrompt
echo       Word closed.
exit /b 0

:CheckVersionStartup
set "CF_VER=%~1"
set "CF_IS_ACTIVE="
if not defined ACTIVE_VER (
    set "ACTIVE_VER=%CF_VER%"
    set "CF_IS_ACTIVE=1"
)
set "CF_CAND="
for /f "tokens=2,*" %%A in ('reg query "HKCU\Software\Microsoft\Office\%CF_VER%\Word\Options" /v STARTUP-PATH 2^>nul ^| find /I "STARTUP-PATH"') do (
    call set "CF_CAND=%%B"
)
if not defined CF_CAND exit /b 0
set "CF_CAND=%CF_CAND:"=%"
if "%CF_CAND%"=="" exit /b 0
if "%CF_CAND:~-1%"=="\" set "CF_CAND=%CF_CAND:~0,-1%"
if "%CF_CAND%"=="" exit /b 0
call :CleanDir "%CF_CAND%"
if not defined CF_IS_ACTIVE exit /b 0
if defined REG_STARTUP exit /b 0
if "%CF_CAND:~1,1%"==":" (
    if not exist "%CF_CAND:~0,3%" exit /b 0
)
set "REG_STARTUP=%CF_CAND%"
set "TARGET_STARTUP=%CF_CAND%"
exit /b 0

:CleanDir
set "CF_DIR=%~1"
if "%CF_DIR%"=="" exit /b 0
if exist "%CF_DIR%\ChatFormatter*.dotm" del /F /Q /A "%CF_DIR%\ChatFormatter*.dotm" >nul 2>&1
if exist "%CF_DIR%\~$ChatFormatter*.dotm" del /F /Q /A "%CF_DIR%\~$ChatFormatter*.dotm" >nul 2>&1
exit /b 0

:RegisterTrust
set "SECKEY=HKCU\Software\Microsoft\Office\%~1\Word\Security\Trusted Locations"
reg query "%SECKEY%" >nul 2>&1
if errorlevel 1 reg add "%SECKEY%" /f >nul 2>&1
set "CF_TPATH=%~2"
if "%CF_TPATH:~0,2%"=="\\" (
    reg add "%SECKEY%" /v AllowNetworkLocations /t REG_DWORD /d 1 /f >nul 2>&1
)
reg add "%SECKEY%\LocationChatFormatter" /v Path /t REG_EXPAND_SZ /d "%~2" /f >nul 2>&1
reg add "%SECKEY%\LocationChatFormatter" /v Description /t REG_SZ /d "Chat Formatter for Word" /f >nul 2>&1
reg add "%SECKEY%\LocationChatFormatter" /v AllowSubfolders /t REG_DWORD /d 1 /f >nul 2>&1
if not errorlevel 1 set "REG_OK=1"
exit /b 0

:Cleanup
if not defined CF_NO_PAUSE pause
endlocal & exit /b %CF_EXIT%
