# Compiles src\ChatFormatter.bas inside a throwaway document and reports the
# exact compile error, if any.
#
# The add-in route hides every module, so Word's dialog only ever says
# "Compile error in hidden module: ChatFormatter" with no line number.
# Importing the same code into a normal document makes the module visible,
# which is the only way to get a useful diagnostic.

param(
    [string]$Source = 'ChatFormatter.bas'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$bas = Join-Path (Join-Path $root 'src') $Source

if (-not (Test-Path -LiteralPath $bas)) {
    throw "source not found: $bas"
}

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

    // Returns the dialog text blocks found, and dismisses each with OK.
    public static List<string> CollectAndDismiss() {
        var r = new List<string>();
        var handles = new List<IntPtr>();
        EnumWindows((h, l) => {
            if (!IsWord(h) || !IsWindowVisible(h)) return true;
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
            SendMessage(h, 0x0111, (IntPtr)1, IntPtr.Zero); // WM_COMMAND IDOK
        }
        return r;
    }
}
'@

# The installed global template also defines a module called ChatFormatter,
# and "Compile" walks every loaded project. Move it out of STARTUP for the
# duration of the check, otherwise errors always report the hidden add-in
# copy instead of the visible module we just imported.
$startup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
$stashed = @()
Get-Process WINWORD -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
Start-Sleep -Milliseconds 800
Get-ChildItem -LiteralPath $startup -Filter '*.dotm' -EA SilentlyContinue | ForEach-Object {
    $dest = Join-Path $env:TEMP ("cfsv-" + $_.Name)
    Move-Item -LiteralPath $_.FullName -Destination $dest -Force -EA SilentlyContinue
    if (Test-Path -LiteralPath $dest) {
        $stashed += [pscustomobject]@{ From = $_.FullName; To = $dest }
    }
}

# Word refuses to build a VBA project unless this registry value is on.
$vbom = 'HKCU:\Software\Microsoft\Office\16.0\Word\Security'
$hadAccess = $null
$prev = (Get-ItemProperty -Path $vbom -Name AccessVBOM -EA SilentlyContinue).AccessVBOM
if ($null -eq $prev) { $hadAccess = $false } else { $hadAccess = $true }
Set-ItemProperty -Path $vbom -Name AccessVBOM -Value 1 -Type DWord

$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $doc = $word.Documents.Add()
    $vbp = $doc.VBProject
    $vbp.VBComponents.Import($bas)

    # Drop the module-level VB_Name attribute line; Import rejects it.
    $comp = $vbp.VBComponents.Item(1)
    $cm = $comp.CodeModule
    $header = $cm.Lines(1, 1)
    if ($header -match 'Attribute VB_Name') {
        $cm.DeleteLines(1, 1)
    }

    # Trigger a compile. 578 is the VBE "Compile <project>" command id.
    $found = $null
    try {
        $word.VBE.CommandBars.FindControl(1, 578).Execute()
    } catch {
        $found = @("execute failed: $($_.Exception.Message)")
    }

    $found = [WordDialogs]::CollectAndDismiss()

    if ($found) {
        Write-Host 'COMPILE ERRORS:'
        $found | ForEach-Object { Write-Host "  $_" }
        exit 1
    }

    Write-Host 'COMPILE OK'
}
finally {
    if ($word) {
        try { $word.Quit(0) } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    Get-Process WINWORD -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue

    foreach ($s in $stashed) {
        Move-Item -LiteralPath $s.To -Destination $s.From -Force -EA SilentlyContinue
    }

    if ($hadAccess) {
        Set-ItemProperty -Path $vbom -Name AccessVBOM -Value $prev -Type DWord
    } else {
        Remove-ItemProperty -Path $vbom -Name AccessVBOM -EA SilentlyContinue
    }
}
