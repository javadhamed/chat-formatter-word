<#
    Build-time compile gate.

    Loading the built template into Word and calling a macro is what actually
    compiles the module: Word fully compiles a module that comes from a global
    template, which is how errors in procedures that the smoke run never
    reaches still get caught.

    The important detail is who notices the error. A compile error comes up as
    a modal VBA dialog while the macro call is still blocked, so a watcher that
    only polls after the call returns sees nothing. This probe therefore runs a
    background watcher for the whole call, records every dialog Word raises, and
    dismisses it so nothing can wedge the build. Any recorded dialog is a
    failure: a clean run produces no dialogs at all.

    Writes "<STATUS>`r`n<detail>" to -ResultFile: OK, ERROR, or FAILED.
#>
param(
    [Parameter(Mandatory = $true)][string]$Dotm,
    [Parameter(Mandatory = $true)][string]$ResultFile,
    [int]$DialogWaitSec = 25
)

$dialogLog = Join-Path $env:TEMP 'cf_probe_dialogs.txt'
Remove-Item $dialogLog -Force -EA SilentlyContinue

Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Collections.Generic;
public class CP {
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc cb, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr p, EnumWindowsProc cb, IntPtr l);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] static extern IntPtr PostMessage(IntPtr h, uint m, IntPtr wp, IntPtr lp);
    delegate bool EnumWindowsProc(IntPtr h, IntPtr l);

    static bool IsWord(IntPtr h) {
        uint pid; GetWindowThreadProcessId(h, out pid);
        try { return System.Diagnostics.Process.GetProcessById((int)pid).ProcessName.ToUpper() == "WINWORD"; }
        catch { return false; }
    }

    static List<IntPtr> Dialogs() {
        var d = new List<IntPtr>();
        EnumWindows((h, l) => {
            if (!IsWord(h)) return true;
            var c = new StringBuilder(256); GetClassName(h, c, 256);
            if (c.ToString() == "#32770") d.Add(h);
            return true;
        }, IntPtr.Zero);
        return d;
    }

    // Records the text of every Word dialog and closes it, so the caller never
    // blocks on a modal message but still learns what it said.
    public static List<string> Sweep() {
        var r = new List<string>();
        foreach (var h in Dialogs()) {
            var parts = new List<string>();
            EnumChildWindows(h, (c, l) => {
                var t = new StringBuilder(1024); GetWindowText(c, t, 1024);
                if (t.Length > 0) parts.Add(t.ToString());
                return true;
            }, IntPtr.Zero);
            if (parts.Count > 0) r.Add(string.Join(" ", parts.ToArray()));
            PostMessage(h, 0x0010, IntPtr.Zero, IntPtr.Zero); // WM_CLOSE
        }
        return r;
    }
}
'@

$securityKey = 'HKCU:\Software\Microsoft\Office\16.0\Word\Security'
$origAccessVBOM = $null
try { $origAccessVBOM = (Get-ItemProperty -Path $securityKey -Name AccessVBOM -EA SilentlyContinue).AccessVBOM } catch { }

$startup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'

function Save-ProbeResult([string]$status, [string]$detail) {
    ($status + "`r`n" + $detail) | Set-Content -LiteralPath $ResultFile -Encoding UTF8
}

$stashed = @()
$word = $null
$doc = $null

try {
    # Keep every other template out of STARTUP: a second one exporting the
    # same public names would make the macro call ambiguous.
    Get-ChildItem -LiteralPath $startup -Filter '*.dotm' -EA SilentlyContinue | ForEach-Object {
        $dest = Join-Path $env:TEMP ('cf_stash_' + $_.Name)
        Move-Item -LiteralPath $_.FullName -Destination $dest -Force -EA SilentlyContinue
        $stashed += [pscustomobject]@{ From = $_.FullName; To = $dest }
    }

    New-ItemProperty -Path $securityKey -Name AccessVBOM -Value 1 -PropertyType DWord -Force -EA SilentlyContinue | Out-Null

    Copy-Item -LiteralPath $Dotm -Destination (Join-Path $startup 'ChatFormatter.dotm') -Force
    Get-Process WINWORD -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
    Start-Sleep -Milliseconds 800

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add()

    # Watcher first, then the macro call. Anything Word raises while the call
    # is in flight -- including the "Compile error in hidden module" dialog that
    # only ever appears in this situation -- lands in the log.
    $stop = Join-Path $env:TEMP 'cf_probe_stop.flag'
    Remove-Item $stop -Force -EA SilentlyContinue
    $watcher = Start-Job -ScriptBlock {
        param($log, $stopFlag)
        Add-Type -TypeDefinition @"
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public class WD {
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc cb, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr p, EnumWindowsProc cb, IntPtr l);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] static extern IntPtr PostMessage(IntPtr h, uint m, IntPtr wp, IntPtr lp);
    delegate bool EnumWindowsProc(IntPtr h, IntPtr l);
    static bool IsWord(IntPtr h) {
        uint pid; GetWindowThreadProcessId(h, out pid);
        try { return System.Diagnostics.Process.GetProcessById((int)pid).ProcessName.ToUpper() == "WINWORD"; }
        catch { return false; }
    }
    public static List<string> Sweep() {
        var d = new List<IntPtr>();
        EnumWindows((h, l) => {
            if (!IsWord(h)) return true;
            var c = new StringBuilder(256); GetClassName(h, c, 256);
            if (c.ToString() == "#32770") d.Add(h);
            return true;
        }, IntPtr.Zero);
        var r = new List<string>();
        foreach (var h in d) {
            var parts = new List<string>();
            EnumChildWindows(h, (c, l) => {
                var t = new StringBuilder(1024); GetWindowText(c, t, 1024);
                if (t.Length > 0) parts.Add(t.ToString());
                return true;
            }, IntPtr.Zero);
            if (parts.Count > 0) r.Add(string.Join(" ", parts.ToArray()));
            PostMessage(h, 0x0010, IntPtr.Zero, IntPtr.Zero);
        }
        return r;
    }
}
"@
        $seen = @{}
        while (-not (Test-Path -LiteralPath $stopFlag)) {
            foreach ($t in [WD]::Sweep()) {
                if ($t.Trim().Length -gt 0 -and -not $seen.ContainsKey($t)) {
                    $seen[$t] = $true
                    Add-Content -LiteralPath $log -Value $t
                }
            }
            Start-Sleep -Milliseconds 250
        }
    } -ArgumentList $dialogLog, $stop

    $runError = $null
    try {
        [void]$word.Run('CF_CompileProbe')
    }
    catch {
        $runError = $_.Exception.Message
    }

    # Give a late dialog a moment, then stop the watcher.
    Start-Sleep -Seconds 2
    New-Item -ItemType File -Path $stop -Force | Out-Null
    $null = Wait-Job $watcher -Timeout 10
    Receive-Job $watcher -EA SilentlyContinue | Out-Null
    Remove-Job $watcher -Force -EA SilentlyContinue
    [CP]::Sweep() | Out-Null

    $dialogs = @()
    if (Test-Path -LiteralPath $dialogLog) {
        $dialogs = @(Get-Content -LiteralPath $dialogLog | Where-Object { $_.Trim().Length -gt 0 })
    }

    if ($dialogs.Count -gt 0) {
        Save-ProbeResult 'ERROR' (($dialogs | Select-Object -Unique) -join ' | ')
        throw 'probe-failed'
    }

    if ($runError) {
        Save-ProbeResult 'ERROR' ('Run(CF_CompileProbe) failed: ' + $runError)
        throw 'probe-failed'
    }

    Save-ProbeResult 'OK' 'Template loaded and CF_CompileProbe ran the whole pipeline; Word raised no dialog.'
}
catch {
    if ($_.Exception.Message -ne 'probe-failed') {
        Save-ProbeResult 'FAILED' $_.Exception.Message
    }
}
finally {
    Remove-Item $stop -Force -EA SilentlyContinue
    Get-Process WINWORD -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
    Get-ChildItem -LiteralPath $startup -Filter '*.dotm' -EA SilentlyContinue |
        Remove-Item -Force -EA SilentlyContinue
    foreach ($s in $stashed) {
        Move-Item -LiteralPath $s.To -Destination $s.From -Force -EA SilentlyContinue
    }
    if ($null -eq $origAccessVBOM) {
        Remove-ItemProperty -Path $securityKey -Name AccessVBOM -EA SilentlyContinue
    }
    else {
        Set-ItemProperty -Path $securityKey -Name AccessVBOM -Value $origAccessVBOM -Type DWord
    }
}