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
set "CF_SIZE=37410"
set "CF_SHA256=a435f6d0ef027dff6f2fc198ca9ba9d8ef820f7dbfdad47e7961ae1757f28c76"
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
echo     3. Paste it into Word
echo     4. Open the Add-ins tab and click "Format Chat Text"
echo.
echo   Other macros ^(Alt+F8^): FormatSelection, ShowSettings
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
AAAAAA4ADgCMAwAAgI4AAAAA
-----END CHATFORMATTER DOTM-----
