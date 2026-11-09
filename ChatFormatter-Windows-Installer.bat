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
set "CF_SIZE=30784"
set "CF_SHA256=ed0807813aded4bbf216c16c5c8cec69f583a2ed5d02ec147a0bfc74a9f3e03a"
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
5c1LtAwfn9BCyEQnYywd9X94bdtVNVgOpt+BRLJ1+JV3IIEuzqVfTzLcmFNUZnmxsXnCgBX2R5iZ
R56guNTiQV0/L9zfC8e0cnFLBxyrmkTMY/vf7Vx9cBvFFX97Okm2IsWy4jh2cOKz5DgylpS7k+zY
CUr0ebaJjVzbSZpicGTpRA4rkiLJ5IswZzsfQEK+SgstlJiB0kwopQE6nQ4zjZ2kQGcYvqfTaWea
BJgO/Q9K+eqQSd/KSup2Co0Df3SmvPFvtJb2dm933+17v9vdd4yrhVOJj3bIUrkEZnvji9vfi0Ud
OXkXEYmXv+4OhR5UqrdbOyIl3WGrfhed75sFptHH8PNucWnW9kG4t6c3InV26Y9pjml+oOnv7gl1
6nqpa07Pt4aqlGz9512fIA8XWM5vXlh4ezhUHyAOTtjC/mWgB6L5TX103e9gNMOtX+bRoY1+QnPx
Aa1zI9vWU9HzQqWq+avmlYZjuu/pfk6sLiH0t8AgfzC9nIUI28Hm8nrXoQlX/GjZQ7fXtcv5s4Yj
7+4blre3cw4uSLbF5cR7D79qm9jjJ51cKi/fJlfdo+RqIhl+YfW+H5Okklr5p4YtJMObtpPVPWlu
1BxBDw0dtIp7nKLG63A9RhIZImXvlGWBPM6Wl1o0owc6X0tlRvIf6Gy7d9TP0X/27T/XP3p3OCLZ
E4fOagO1W4x9rC3hGDIpJw51pg4b7VGi/tS72HrE/t2mPy46zPRnlU8319uDDR8yv9acbBCaVvsX
rmYJfFJjNZZru5VUvEzLaiemGl8yaiH4eTJtblpSXgqn9y1fcoKce+/gRfiJ8KT2d5XS4bg67+ID
i1akjw8fN897S5P17dReWpY/df8Ji9l1SWx6fpJ5Wrz4M/EOHn70aplYJs4V9/OrHBeUD1atdp/y
jJBh+WU5c+rWIyS1ZoKMkoqutEKUZWQkE9eDefBM9/x2JOVrLzVM1LZVto0YP+p/lPmw/wicElwn
4cmuDeviG1+/1f4ce6z8nffhzZ2MeDbaPbyRHJuv/0Xps2U88w70yk9Y3hbOn+NaPKruad3p0kvn
VtZKbz0/KG2QjkkJck7SGFfUth84aZZNBq6EHqZGVxsNhj2JXF92cFavtREKu6fRYOhScx5nnyu8
QJu0U0+unk3MvQucnNDYyHm5bTjqFQfRCeHZ+gYbWQ7aO6LJeU0jGUpY5Di4+rMvjBinCNDTAQ+V
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
AAAAAA==
-----END CHATFORMATTER DOTM-----
