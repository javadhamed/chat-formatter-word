<#
.SYNOPSIS
    Builds ChatFormatter.dotm from the VBA sources in src/.

.DESCRIPTION
    Requires Microsoft Word (any recent version) and PowerShell.
    Uses Word COM automation, so Word must be able to run headlessly on
    this machine. Temporarily enables "Trust access to the VBA project
    object model" (AccessVBOM), which Word requires before any tool may
    write code into a document, then restores the previous value.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\build.ps1
#>

[CmdletBinding()]
param(
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

# Shared helper for reading (and dismissing) modal Word dialogs.
Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Collections.Generic;
public class WordDialogs {
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc cb, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr p, EnumWindowsProc cb, IntPtr l);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, uint msg, IntPtr wp, IntPtr lp);
    delegate bool EnumWindowsProc(IntPtr h, IntPtr l);

    static bool IsWord(IntPtr h) {
        uint pid; GetWindowThreadProcessId(h, out pid);
        try { return System.Diagnostics.Process.GetProcessById((int)pid).ProcessName.ToUpper() == "WINWORD"; }
        catch { return false; }
    }

    // BM_CLICK needs zero wParam/lParam; passing 1 silently does nothing,
    // which leaves the dialog on screen and lets a broken build report OK.
    const uint BM_CLICK = 0x00F5;
    const uint WM_CLOSE = 0x0010;

    public static List<string> CollectAndDismiss() {
        var r = new List<string>();
        var handles = new List<IntPtr>();
        EnumWindows((h, l) => {
            if (!IsWord(h)) return true;
            var c = new StringBuilder(256); GetClassName(h, c, 256);
            if (c.ToString() == "#32770") handles.Add(h);
            return true;
        }, IntPtr.Zero);
        foreach (var h in handles) {
            EnumChildWindows(h, (ch, cl) => {
                var ct = new StringBuilder(2048); GetWindowText(ch, ct, 2048);
                if (ct.Length > 0) r.Add(ct.ToString());
                return true;
            }, IntPtr.Zero);
            SendMessage(h, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
            // If a click did not take it down, close it outright so the build
            // never leaves a modal dialog blocking the next Word launch.
            if (IsWindowVisible(h)) SendMessage(h, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        }
        return r;
    }
}
'@

$root       = Split-Path -Parent $PSScriptRoot
$srcDir     = Join-Path $root 'src'
$basFile    = Join-Path $srcDir 'ChatFormatter.bas'
$clsFile    = Join-Path $srcDir 'ThisDocument.cls'

if (-not $OutputPath) { $OutputPath = Join-Path $root 'ChatFormatter.dotm' }

foreach ($f in @($basFile, $clsFile)) {
    if (-not (Test-Path -LiteralPath $f)) { throw "Missing source: $f" }
}

# ---------------------------------------------------------------------------
# Structural pre-check
#
# Word's own compile is the only real authority, but it reports syntax
# failures as a modal dialog with no line number, and that dialog is easy
# for the compile step below to miss -- which lets a template that cannot
# compile be saved and installed. This catches the common case (a block
# keyword sitting on a line after "Then <statement>", which VBA reads as
# "Else without If") and names the line before Word ever runs.
# ---------------------------------------------------------------------------
$structErrors = @()
foreach ($f in @($basFile, $clsFile)) {
    $raw = [System.IO.File]::ReadAllLines($f, [System.Text.Encoding]::UTF8)

    # Join VBA line-continuation (" _") first. A condition split across
    # physical lines only exposes its "Then" on the final one, so without
    # this the block opener is missed and the matching "End If" looks stray.
    $logical = New-Object System.Collections.Generic.List[object]
    $buf = ''
    $bufLine = 0
    for ($i = 0; $i -lt $raw.Length; $i++) {
        $t = $raw[$i].Trim()
        if ($buf -eq '') { $bufLine = $i + 1 }
        if ($t -match '_\s*$') {
            $buf += ($t -replace '_\s*$', '') + ' '
            continue
        }
        $logical.Add([pscustomobject]@{ Line = $bufLine; Code = ($buf + $t) })
        $buf = ''
    }
    if ($buf -ne '') { $logical.Add([pscustomobject]@{ Line = $bufLine; Code = $buf }) }

    $depth = 0
    foreach ($entry in $logical) {
        $code = $entry.Code.Trim()
        if ($code -eq '' -or $code.StartsWith("'")) { continue }
        $code = ($code -split "'", 2)[0].Trim()
        if ($code -eq '') { continue }

        # "If ... Then <something>" is a single-line If: it opens no block,
        # so a following Else/End If is a structural error.
        $opensBlock = $false
        if ($code -match '^(?i)If\b') {
            if ($code -match '(?i)\bThen\b(.*)$') {
                if ($Matches[1].Trim() -eq '') { $opensBlock = $true }
            }
        }
        elseif ($code -match '^(?i)ElseIf\b') {
            if ($depth -eq 0) {
                $structErrors += ("{0}:{1}: 'ElseIf' with no open 'If'" -f `
                    (Split-Path -Leaf $f), $entry.Line)
            }
        }
        elseif ($code -match '^(?i)Else\b') {
            if ($depth -eq 0) {
                $structErrors += ("{0}:{1}: 'Else' with no open 'If' -> {2}" -f `
                    (Split-Path -Leaf $f), $entry.Line, $code)
            }
        }
        elseif ($code -match '^(?i)End\s+If\b') {
            if ($depth -eq 0) {
                $structErrors += ("{0}:{1}: 'End If' with no open 'If' -> {2}" -f `
                    (Split-Path -Leaf $f), $entry.Line, $code)
            }
            else { $depth-- }
        }

        if ($opensBlock) { $depth++ }
    }
    if ($depth -ne 0) {
        $structErrors += ("{0}: {1} unclosed block(s) at end of file" -f `
            (Split-Path -Leaf $f), $depth)
    }
}
if ($structErrors) {
    Write-Host ''
    Write-Host 'STRUCTURAL ERRORS (VBA would report these as compile errors):'
    $structErrors | ForEach-Object { Write-Host "  $_" }
    throw 'Source failed the structural pre-check; refusing to build.'
}
Write-Host 'Structure: OK'

# ---------------------------------------------------------------------------
# Trust access to the VBA project object model - required by Word COM
# ---------------------------------------------------------------------------
$securityKey = 'HKCU:\Software\Microsoft\Office\16.0\Word\Security'
if (-not (Test-Path -LiteralPath $securityKey)) {
    New-Item -Path $securityKey -Force | Out-Null
}
$origAccessVBOM = (Get-ItemProperty -Path $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
Write-Host "AccessVBOM was: $origAccessVBOM"

# Declared outside the try so the finally block can always restore it.
$stashed = @()

try {
    Set-ItemProperty -Path $securityKey -Name AccessVBOM -Value 1 -Type DWord

    # -----------------------------------------------------------------------
    # Create the template
    # -----------------------------------------------------------------------
    # Word's compile runs across every loaded project, including global
    # templates from STARTUP. An older broken copy of this add-in sitting
    # there would be reported as "compile error in hidden module:
    # ChatFormatter" and mask the real result, so move it aside first.
    $startup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
    $stashed = @()
    Get-Process WINWORD -ErrorAction SilentlyContinue |
        Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 800
    Get-ChildItem -LiteralPath $startup -Filter '*.dotm' -ErrorAction SilentlyContinue |
        ForEach-Object {
            $dest = Join-Path $env:Temp ("cfbuild-" + $_.Name)
            Move-Item -LiteralPath $_.FullName -Destination $dest -Force -ErrorAction SilentlyContinue
            if (Test-Path -LiteralPath $dest) {
                $stashed += [pscustomobject]@{ From = $_.FullName; To = $dest }
            }
        }

    Write-Host "Launching Word..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    try {
        $doc = $word.Documents.Add()

        # 15 = wdFormatXMLTemplateMacroEnabled (.dotm)
        $doc.SaveAs2([ref]$OutputPath, [ref]15)
        Write-Host "Created template: $OutputPath"

        $vbProj = $doc.VBProject

        # Wipe default modules, keep ThisDocument.
        foreach ($comp in @($vbProj.VBComponents)) {
            if ($comp.Name -ne 'ThisDocument') {
                $vbProj.VBComponents.Remove($comp)
            }
        }

        # ---- standard module -------------------------------------------------
        # Import, not CodeModule.AddFromString. AddFromString silently drops
        # code on a module this size, and the template then reloads with
        # "Compile error in hidden module: ChatFormatter" even though the
        # in-session compile looked clean.
        $std = $vbProj.VBComponents.Import($basFile)
        $std.Name = 'ChatFormatter'
        Write-Host "Injected module: ChatFormatter ($($std.CodeModule.CountOfLines) lines)"

        # ---- ThisDocument event code ----------------------------------------
        $cls = [System.IO.File]::ReadAllText($clsFile, [System.Text.Encoding]::UTF8)
        $cls = $cls -replace '(?s)^.*?Attribute VB_Customizable = True\r?\n', ''
        $cls = $cls -replace '(?m)^Attribute .*\r?\n', ''

        $thisDoc = $vbProj.VBComponents.Item('ThisDocument')
        if ($env:CF_SKIP_THISDOCUMENT -eq '1') {
            Write-Host "Injected events: SKIPPED (diagnostic)"
        }
        else {
            $thisDoc.CodeModule.AddFromString($cls)
            Write-Host "Injected events: ThisDocument ($($thisDoc.CodeModule.CountOfLines) lines)"
        }

        # No compile attempt here. The VBE Compile command reports
        # Enabled = $true but silently does nothing under automation, which
        # makes it worse than useless as a gate. The real check runs after the
        # save, by loading the template and calling a macro.

        $doc.Save()
        Write-Host "Saved."
    }
    finally {
        $doc.Close([ref]0)      # 0 = wdDoNotSaveChanges
        $word.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
}
finally {
    foreach ($s in $stashed) {
        Move-Item -LiteralPath $s.To -Destination $s.From -Force -ErrorAction SilentlyContinue
    }

    # Restore the original VBA trust setting.
    if ($null -eq $origAccessVBOM) {
        Remove-ItemProperty -Path $securityKey -Name AccessVBOM -ErrorAction SilentlyContinue
    }
    else {
        Set-ItemProperty -Path $securityKey -Name AccessVBOM -Value $origAccessVBOM -Type DWord
    }
    Write-Host "AccessVBOM restored to: $origAccessVBOM"
}

# ---------------------------------------------------------------------------
# Compile probe
#
# The VBE "Compile" command cannot be trusted as a gate. FindControl(1,578)
# can return without compiling anything, and the dialog it raises is created
# asynchronously, so sampling dialogs straight after Execute() races it. That
# combination reported "Compile: OK" for a template that Word then refused to
# run, which is worse than no check at all.
#
# Instead, load the built template into its own Word and call a macro: Run()
# compiles the project for real. The probe runs as a separate, time boxed
# process, so the modal "Compile error" dialog can never wedge the build and a
# missing answer is itself a failure.
# ---------------------------------------------------------------------------
$probeResult = Join-Path $env:TEMP 'cf_compile_probe.txt'
Remove-Item -LiteralPath $probeResult -Force -ErrorAction SilentlyContinue

$p = Start-Process powershell.exe -PassThru -NoNewWindow `
    -ArgumentList @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass',
        '-File', ('"' + (Join-Path $PSScriptRoot 'compile-probe.ps1') + '"'),
        ('"' + $OutputPath + '"'),
        ('"' + $probeResult + '"')
    ) `
    -RedirectStandardOutput (Join-Path $env:TEMP 'cf_probe_out.txt') `
    -RedirectStandardError  (Join-Path $env:TEMP 'cf_probe_err.txt')

$probeOk = $false
$probeMsgs = @()
$deadline = (Get-Date).AddSeconds(120)

function Read-ProbeVerdict {
    if (-not (Test-Path -LiteralPath $probeResult)) { return '' }
    # ReadAllText, not Get-Content: the probe writes UTF8 with a BOM, so a
    # regex anchored at ^OK would never match a Get-Content result.
    try { return [System.IO.File]::ReadAllText($probeResult) } catch { return '' }
}

$probeOk = $false
$probeMsgs = @()
$deadline = (Get-Date).AddSeconds(120)

while ((Get-Date) -lt $deadline) {
    $verdict = Read-ProbeVerdict
    if ($verdict -match '(?s)^\s*OK\b') { $probeOk = $true; break }
    if ($verdict -match '(?s)^\s*(ERROR|FAILED)\b') { break }

    # Dismiss whatever Word raises so the probe can never be wedged, but do
    # not treat it as a verdict: the probe interprets its own dialogs and
    # unrelated ones (activation, recovery) would otherwise fail a good build.
    $dl = [WordDialogs]::CollectAndDismiss()
    if ($dl) { $probeMsgs += $dl }

    if ($p.HasExited) {
        # The child writes its verdict just before exiting, so a read taken a
        # moment ago can be empty or half written. Settle, then trust the file.
        Start-Sleep -Milliseconds 400
        if ((Read-ProbeVerdict) -match '(?s)^\s*OK\b') { $probeOk = $true }
        break
    }
    Start-Sleep -Milliseconds 500
}

if (-not $p.HasExited) { try { $p.Kill() } catch { } }
Start-Sleep -Milliseconds 600
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

if ($probeOk) {
    Write-Host 'Compile: OK'
}
else {
    Write-Host ''
    Write-Host 'COMPILE ERRORS:'
    if ($probeMsgs) {
        $probeMsgs | ForEach-Object { Write-Host "  $_" }
    }
    elseif (Test-Path -LiteralPath $probeResult) {
        [System.IO.File]::ReadAllText($probeResult).Trim() -split "`r?`n" |
            ForEach-Object { Write-Host "  $_" }
    }
    else {
        Write-Host '  compile probe produced no result (timed out after 120s)'
    }
    throw 'ChatFormatter failed to compile; template not usable.'
}

Write-Host ''
Write-Host "Build complete: $OutputPath"
Write-Host "Next: run install.bat to deploy it into Word's STARTUP folder."