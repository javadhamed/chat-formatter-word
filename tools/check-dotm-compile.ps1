# Compile the actual built .dotm template and report any errors.
# This loads the real template file, so any reference/template-specific
# issues will surface.

param(
    [string]$DotmPath = 'ChatFormatter.dotm'
)

$ErrorActionPreference = 'Stop'

Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Collections.Generic;
public class WDlg {
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
            SendMessage(h, 0x0111, (IntPtr)1, IntPtr.Zero);
        }
        return r;
    }
}
'@

$dotm = (Resolve-Path -LiteralPath $DotmPath).Path
if (-not (Test-Path -LiteralPath $dotm)) { throw "Not found: $dotm" }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    # Open the template as a document to make its modules visible
    $doc = $word.Documents.Open($dotm, $false, $true, $true)  # readonly, visible, open as template
    $vbp = $doc.VBProject

    # Find the ChatFormatter module
    $mod = $null
    foreach ($c in $vbp.VBComponents) {
        if ($c.Name -eq 'ChatFormatter') { $mod = $c; break }
    }
    if ($null -eq $mod) { throw "ChatFormatter module not found in $dotm" }

    # Delete any VB_Name attribute line
    $cm = $mod.CodeModule
    if ($cm.Lines(1, 1) -match 'Attribute VB_Name') { $cm.DeleteLines(1, 1) }

    # Force compile
    try {
        $word.VBE.CommandBars.FindControl(1, 578).Execute()
    }
    catch {
        # ignore - dialog will be collected
    }

    Start-Sleep -Milliseconds 500
    $dl = [WDlg]::CollectAndDismiss()
    if ($dl) {
        Write-Host "COMPILE ERRORS in ${dotm}:"
        $dl | ForEach-Object { Write-Host "  $_" }
        exit 1
    }

    Write-Host "COMPILE OK for ${dotm}"
}
finally {
    try { $word.Quit(0) } catch { }
    Get-Process WINWORD -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
}