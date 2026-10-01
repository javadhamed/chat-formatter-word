# Runs the formatter one stage at a time and reports which stage wedges Word.
#
# tools\test.ps1 drives the whole pipeline in a single COM call, so a hang in
# any stage looks exactly like a hang in any other. This runs the macro's
# CF_TestStage entry point one stage at a time from a background job, writing
# progress to a log file, while the foreground polls that file and watches for
# modal dialogs. The last stage that logged BEGIN without a matching OK is the
# one that hung.

param(
    [int]$StageTimeout = 12,
    [int]$LastStage = 14
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

$logPath = Join-Path $env:TEMP 'cfstages.log'

$bt = [string][char]96
$fence = $bt + $bt + $bt
$tick2 = $bt + $bt

$sample = @(
    '# Heading One'
    "Some **bold** and *italic* and ~~strike~~ and $tick2 code $tick2 here."
    '## Heading Two'
    '> a quoted line'
    '- first bullet'
    '1. one'
    '| Name | Age |'
    '|------|-----|'
    '| Ali  | 30  |'
    '| Sara | 25  |'
    $fence
    'plain code line'
    $fence
) -join "`r"

$startup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
$dotm = Join-Path (Split-Path -Parent $PSScriptRoot) 'ChatFormatter.dotm'
Get-Process WINWORD -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
Start-Sleep -Milliseconds 800
Copy-Item -LiteralPath $dotm -Destination $startup -Force
Remove-Item -LiteralPath $logPath -Force -ErrorAction SilentlyContinue

$sec = 'HKCU:\Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\LocationChatFormatter'
if (-not (Test-Path -LiteralPath $sec)) { New-Item -Path $sec -Force | Out-Null }
Set-ItemProperty -Path $sec -Name Path -Value $startup
Set-ItemProperty -Path $sec -Name AllowSubfolders -Value 1 -Type DWord
Set-ItemProperty -Path 'HKCU:\Software\ChatFormatter' -Name AutoFormat -Value '0' -EA SilentlyContinue

# All Word interaction happens in a job; the foreground stays responsive so it
# can keep watching for dialogs while a stage is blocked.
$job = Start-Job -ArgumentList $sample, $LastStage, $logPath -ScriptBlock {
    param($sample, $lastStage, $logPath)

    function Say($msg) { Add-Content -LiteralPath $logPath -Value $msg -Encoding utf8 }

    $names = @{
        1 = 'PipeTables'; 2 = 'TSVTables'; 3 = 'HorizontalRules'; 4 = 'CodeBlocks'
        5 = 'Headings'; 6 = 'Bold'; 7 = 'Italic'; 8 = 'Strike'; 9 = 'InlineCode'
        10 = 'Links'; 11 = 'Blockquotes'; 12 = 'Bullets'; 13 = 'Numbered'
        14 = 'StrayMarkers'
    }

    Say "JOBSTART"
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add()
    $doc.Content.Text = $sample
    Say "PARAS $($doc.Paragraphs.Count)"

    foreach ($n in 1..$lastStage) {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        Say "BEGIN $($names[$n])"
        try {
            # Word's Application.Run takes ByRef args, so PowerShell needs
            # an explicit [ref] wrapper here.
            $word.Run('ChatFormatter.CF_TestStage', [ref]$n)
            Say "OK $($names[$n]) $($sw.ElapsedMilliseconds)ms"
        }
        catch {
            Say "THREW $($names[$n]): $($_.Exception.Message)"
        }
    }

    Say 'TEXTSTART'
    Add-Content -LiteralPath $logPath -Value ($doc.Content.Text -replace "`r", "`n") -Encoding utf8
    Say 'TEXTEND'
    Say "TABLES $($doc.Tables.Count)"
    if ($doc.Tables.Count -gt 0) {
        $t = $doc.Tables.Item(1)
        Say "TABLESIZE $($t.Rows.Count)x$($t.Columns.Count)"
        for ($r = 1; $r -le $t.Rows.Count; $r++) {
            $cells = @()
            for ($c = 1; $c -le $t.Columns.Count; $c++) {
                $cells += ($t.Cell($r, $c).Range.Text -replace "[\r\a]", '')
            }
            Say "ROW$r " + ($cells -join ' ~ ')
        }
    }

    $doc.Close([ref]0)
    $word.Quit()
    Say 'FINISHED'
}

$seen = 0
$deadline = (Get-Date).AddSeconds(($StageTimeout * $LastStage) + 60)
$hungStage = ''

while ((Get-Date) -lt $deadline) {
    $state = (Get-Job -Id $job.Id -ErrorAction SilentlyContinue).State
    $lines = @(Get-Content -LiteralPath $logPath -ErrorAction SilentlyContinue)
    if ($lines.Count -gt $seen) {
        for ($i = $seen; $i -lt $lines.Count; $i++) { Write-Host $lines[$i] }
        $seen = $lines.Count

        $begun = @($lines | Where-Object { $_ -like 'BEGIN *' })
        $done = @($lines | Where-Object { $_ -like 'OK *' -or $_ -like 'THREW *' })
        if ($begun.Count -gt $done.Count) {
            $hungStage = $begun[$done.Count].Substring(6)
        }
    }

    $dl = [WDlg]::CollectAndDismiss()
    foreach ($d in $dl) {
        if ($d -eq 'OK' -or $d -eq 'Help') { continue }
        Write-Host "DIALOG: $d"
    }

    if ($state -ne 'Running') { break }
    Start-Sleep -Milliseconds 700
}

$finalLines = @(Get-Content -LiteralPath $logPath -ErrorAction SilentlyContinue)
if ($finalLines.Count -gt $seen) {
    for ($i = $seen; $i -lt $finalLines.Count; $i++) { Write-Host $finalLines[$i] }
}

if ($hungStage) {
    Write-Host ''
    Write-Host "HUNG IN STAGE: $hungStage"
}

Get-Job | Stop-Job -ErrorAction SilentlyContinue
Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue
Get-Process WINWORD -EA SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

Remove-Item -LiteralPath $logPath -Force -ErrorAction SilentlyContinue
