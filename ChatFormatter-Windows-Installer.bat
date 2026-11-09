@echo off
setlocal EnableExtensions
chcp 65001 >nul 2>&1

REM ==============================================================
REM   Chat Formatter for Word - Standalone Windows Installer
REM   Self-contained: embeds ChatFormatter.dotm, removes any old
REM   installation first, then installs the new version.
REM ==============================================================

title Chat Formatter - Install / Upgrade

set "TEMPLATE_NAME=ChatFormatter.dotm"
set "CF_SELF=%~f0"
set "CF_SIZE=32686"
set "CF_SHA256=ab528d3f5eb21506b55b0531871c5cc5706d13bd469de9f2d6ca2e67e7d3db63"
set "CF_TMP_DOTM=%TEMP%\ChatFormatter_new_%RANDOM%_%RANDOM%.dotm"
set "CF_B64_TXT=%TEMP%\ChatFormatter_b64_%RANDOM%_%RANDOM%.txt"
set "CF_EXIT=0"

echo.
echo ==============================================================
echo   Chat Formatter for Word - Standalone Installer (Windows)
echo ==============================================================
echo.

REM --- 1. Extract & verify embedded ChatFormatter.dotm ------------
echo [1/5] Preparing new ChatFormatter package...
call :ExtractPayload
if errorlevel 1 (
    echo [ERROR] Failed to extract or verify the embedded template.
    set "CF_EXIT=1"
    goto :Cleanup
)
echo       OK ^(verified %CF_SIZE% bytes^)

REM --- 2. Ensure Word is closed before touching STARTUP -----------
echo [2/5] Checking if Microsoft Word is running...
call :EnsureWordClosed
if errorlevel 1 (
    set "CF_EXIT=1"
    goto :Cleanup
)

REM --- 3. Remove previous version completely ----------------------
echo [3/5] Removing any previous version...
set "DEFAULT_STARTUP=%APPDATA%\Microsoft\Word\STARTUP"
set "TARGET_STARTUP=%DEFAULT_STARTUP%"
set "REG_STARTUP="

call :CleanDir "%DEFAULT_STARTUP%"
for %%V in (14.0 15.0 16.0) do (
    for /f "tokens=2,*" %%A in ('reg query "HKCU\Software\Microsoft\Office\%%V\Word\Options" /v STARTUP-PATH 2^>nul ^| find /I "STARTUP-PATH"') do (
        call set "REG_STARTUP=%%B"
        call :CleanRegStartup
    )
)
if defined REG_STARTUP (
    if "%REG_STARTUP:~-1%"=="\" set "REG_STARTUP=%REG_STARTUP:~0,-1%"
)
if defined REG_STARTUP (
    if not "%REG_STARTUP%"=="" set "TARGET_STARTUP=%REG_STARTUP%"
)

set "CFHOME=%USERPROFILE%"
if not defined CFHOME set "CFHOME=%HOMEDRIVE%%HOMEPATH%"

call :CleanDir "%LOCALAPPDATA%\Microsoft\Word\STARTUP"
call :CleanDir "%LOCALAPPDATA%\Microsoft\Office\Word\STARTUP"
call :CleanDir "%CFHOME%\Documents\Custom Office Templates"

REM Remove legacy v1.0.0 registry key and stale trusted locations
reg delete "HKCU\Software\ChatFormatter" /f >nul 2>&1
for %%V in (14.0 15.0 16.0) do (
    reg delete "HKCU\Software\Microsoft\Office\%%V\Word\Security\Trusted Locations\LocationChatFormatter" /f >nul 2>&1
)

if exist "%CFHOME%\.chatformatter" (
    rmdir /S /Q "%CFHOME%\.chatformatter" >nul 2>&1
)
echo       Previous files and settings cleaned.

REM --- 4. Install new version -------------------------------------
echo [4/5] Installing new version to Word STARTUP...
if not exist "%TARGET_STARTUP%" mkdir "%TARGET_STARTUP%" 2>nul
copy /Y "%CF_TMP_DOTM%" "%TARGET_STARTUP%\%TEMPLATE_NAME%" >nul
if errorlevel 1 (
    echo [ERROR] Copy failed. Close all Word windows and run this again.
    set "CF_EXIT=1"
    goto :Cleanup
)
if not exist "%TARGET_STARTUP%\%TEMPLATE_NAME%" (
    echo [ERROR] Installed file is missing at %TARGET_STARTUP%\%TEMPLATE_NAME%.
    set "CF_EXIT=1"
    goto :Cleanup
)
echo       OK -^> %TARGET_STARTUP%\%TEMPLATE_NAME%

REM --- 5. Register trusted location and write fresh settings ------
echo [5/5] Registering trusted location and default settings...
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
    echo       [WARN] Could not register Trusted Location automatically.
    echo        In Word: File ^> Options ^> Trust Center ^> Trust Center Settings
    echo        ^> Trusted Locations ^> Add "%TARGET_STARTUP%"
)

if not exist "%CFHOME%\.chatformatter" mkdir "%CFHOME%\.chatformatter" 2>nul
(
    echo AutoFormat = 1
    echo Tables = 1
) > "%CFHOME%\.chatformatter\settings.ini"
echo       Created %CFHOME%\.chatformatter\settings.ini

echo.
echo ==============================================================
echo   Installation complete!
echo ==============================================================
echo.
echo   Installed: %TARGET_STARTUP%\%TEMPLATE_NAME%
echo.
echo   How to use:
echo     1. Open Word
echo     2. Copy text from an AI chat (ChatGPT, DeepSeek, Gemini, ...)
echo     3. Paste into Word, select it, and RIGHT-CLICK
echo.
echo   Manual macros (Alt+F8): FormatChatText, ToggleAutoFormat, ShowSettings
echo.
goto :Cleanup

:ExtractPayload
if exist "%CF_TMP_DOTM%" del /F /Q /A "%CF_TMP_DOTM%" >nul 2>&1
where powershell >nul 2>&1
if not errorlevel 1 (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "$t=[IO.File]::ReadAllText($env:CF_SELF);$b1='-----BEGIN '+'CHATFORMATTER DOTM-----';$e1='-----END '+'CHATFORMATTER DOTM-----';$pat='(?ms)^'+[regex]::Escape($b1)+'\r?\n(.*?)\r?\n'+[regex]::Escape($e1)+'(?:\r?\n|$)';$ms=[regex]::Matches($t,$pat);if($ms.Count -eq 0){exit 2};$b=[Convert]::FromBase64String(($ms[$ms.Count-1].Groups[1].Value -replace '[^A-Za-z0-9+/=]',''));if($b.Length -ne [int]$env:CF_SIZE){exit 3};$h=[BitConverter]::ToString([Security.Cryptography.SHA256]::Create().ComputeHash($b)).Replace('-','').ToLower();if($h -ne $env:CF_SHA256){exit 4};[IO.File]::WriteAllBytes($env:CF_TMP_DOTM,$b)" >nul 2>&1
    if not errorlevel 1 (
        if exist "%CF_TMP_DOTM%" exit /b 0
    )
)

REM Fallback when PowerShell is unavailable or restricted
set "CF_IN_BLK="
> "%CF_B64_TXT%" (
    echo -----BEGIN CERTIFICATE-----
    for /f "usebackq delims=" %%L in ("%CF_SELF%") do (
        if "%%L"=="-----END CHATFORMATTER DOTM-----" set "CF_IN_BLK="
        if defined CF_IN_BLK echo %%L
        if "%%L"=="-----BEGIN CHATFORMATTER DOTM-----" set "CF_IN_BLK=1"
    )
    echo -----END CERTIFICATE-----
)
certutil -f -decode "%CF_B64_TXT%" "%CF_TMP_DOTM%" >nul 2>&1
if errorlevel 1 exit /b 1
set "CF_GOT_SIZE="
for %%F in ("%CF_TMP_DOTM%") do set "CF_GOT_SIZE=%%~zF"
if not "%CF_GOT_SIZE%"=="%CF_SIZE%" exit /b 1
exit /b 0

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
echo       Word is still open (you may have an unsaved document).
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

:CleanRegStartup
if defined REG_STARTUP call :CleanDir "%REG_STARTUP%"
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
reg add "%SECKEY%\LocationChatFormatter" /v Path /t REG_EXPAND_SZ /d "%~2" /f >nul 2>&1
reg add "%SECKEY%\LocationChatFormatter" /v Description /t REG_SZ /d "Chat Formatter for Word" /f >nul 2>&1
reg add "%SECKEY%\LocationChatFormatter" /v AllowSubfolders /t REG_DWORD /d 1 /f >nul 2>&1
if not errorlevel 1 set "REG_OK=1"
exit /b 0

:Cleanup
if exist "%CF_TMP_DOTM%" del /F /Q /A "%CF_TMP_DOTM%" >nul 2>&1
if exist "%CF_B64_TXT%" del /F /Q /A "%CF_B64_TXT%" >nul 2>&1
if not defined CF_NO_PAUSE pause
endlocal & exit /b %CF_EXIT%

-----BEGIN CHATFORMATTER DOTM-----
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
eG1sUEsFBgAAAAAOAA4AjAMAAAx8AAAAAA==
-----END CHATFORMATTER DOTM-----
