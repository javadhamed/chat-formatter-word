@echo off
setlocal EnableExtensions

REM ==============================================================
REM   Chat Formatter for Word - Standalone Windows Installer
REM   Self-contained: embeds ChatFormatter.dotm, removes any old
REM   installation first, then installs the new version.
REM ==============================================================

title Chat Formatter - Install / Upgrade

set "TEMPLATE_NAME=ChatFormatter.dotm"
set "CF_SELF=%~f0"
set "CF_SIZE=39032"
set "CF_SHA256=29b386708c70e2228923b56905a5a307e40c11e7a381642f7f4eb2b2d7d41f6c"
set "CF_TMP_DOTM=%TEMP%\ChatFormatter_new_%RANDOM%_%RANDOM%.dotm"
set "CF_B64_TXT=%TEMP%\ChatFormatter_b64_%RANDOM%_%RANDOM%.txt"
set "CF_EXIT=0"

echo.
echo ==============================================================
echo   Chat Formatter for Word - Standalone Installer (Windows)
echo ==============================================================
echo.
echo   User: %USERNAME%

REM --- 1. Extract & verify embedded ChatFormatter.dotm ------------
echo [1/5] Preparing new ChatFormatter package...
call :ExtractPayload
if errorlevel 1 (
    echo [ERROR] Failed to extract or verify the embedded template.
    echo         Make sure you downloaded the real .bat file ^(not an HTML page^).
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
copy /Y "%CF_TMP_DOTM%" "%TARGET_STARTUP%\%TEMPLATE_NAME%" >nul
if errorlevel 1 (
    echo [ERROR] Copy failed. Close all Word windows and run this again.
    set "CF_EXIT=1"
    goto :Cleanup
)
set "CF_INST_SIZE="
for %%F in ("%TARGET_STARTUP%\%TEMPLATE_NAME%") do set "CF_INST_SIZE=%%~zF"
if not "%CF_INST_SIZE%"=="%CF_SIZE%" (
    echo [ERROR] Installed file verification failed at:
    echo         %TARGET_STARTUP%\%TEMPLATE_NAME%
    set "CF_EXIT=1"
    goto :Cleanup
)
echo       OK -^> %TARGET_STARTUP%\%TEMPLATE_NAME% ^(%CF_INST_SIZE% bytes^)

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
if exist "%CFHOME%\.chatformatter\settings.ini" (
    echo       Created %CFHOME%\.chatformatter\settings.ini
)

echo.
echo ==============================================================
echo   Installation complete!
echo ==============================================================
echo.
echo   Installed: %TARGET_STARTUP%\%TEMPLATE_NAME%
echo.
echo   How to use:
echo     1. Open Word
echo     2. Copy text from an AI chat ^(ChatGPT, DeepSeek, Gemini, ...^)
echo     3. Paste into Word, select it, and RIGHT-CLICK
echo.
echo   Manual macros ^(Alt+F8^): FormatChatText, ToggleAutoFormat, ShowSettings
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
    echo       PowerShell decode skipped; trying certutil fallback...
)

REM Fallback when PowerShell is unavailable or restricted
> "%CF_B64_TXT%" echo -----BEGIN CERTIFICATE-----
findstr /R "^[A-Za-z0-9+/][A-Za-z0-9+/=]*$" "%CF_SELF%" >> "%CF_B64_TXT%"
>> "%CF_B64_TXT%" echo -----END CERTIFICATE-----

certutil -f -decode "%CF_B64_TXT%" "%CF_TMP_DOTM%" >nul 2>&1
if errorlevel 1 exit /b 1
set "CF_GOT_SIZE="
for %%F in ("%CF_TMP_DOTM%") do set "CF_GOT_SIZE=%%~zF"
if not "%CF_GOT_SIZE%"=="%CF_SIZE%" exit /b 1

set "CF_HASH_LINE="
for /f "skip=1 delims=" %%H in ('certutil -hashfile "%CF_TMP_DOTM%" SHA256 2^>nul') do (
    if not defined CF_HASH_LINE set "CF_HASH_LINE=%%H"
)
if not defined CF_HASH_LINE exit /b 0
set "CF_HASH_LINE=%CF_HASH_LINE: =%"
echo "%CF_HASH_LINE%" | findstr /R "^"[0-9A-Fa-f][0-9A-Fa-f]*"$" >nul || exit /b 0
if /I not "%CF_HASH_LINE%"=="%CF_SHA256%" exit /b 1
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
if exist "%CF_TMP_DOTM%" del /F /Q /A "%CF_TMP_DOTM%" >nul 2>&1
if exist "%CF_B64_TXT%" del /F /Q /A "%CF_B64_TXT%" >nul 2>&1
if not defined CF_NO_PAUSE pause
endlocal & exit /b %CF_EXIT%

-----BEGIN CHATFORMATTER DOTM-----
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
-----END CHATFORMATTER DOTM-----
