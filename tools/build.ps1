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

$root       = Split-Path -Parent $PSScriptRoot
$srcDir     = Join-Path $root 'src'
$basFile    = Join-Path $srcDir 'ChatFormatter.bas'
$clsFile    = Join-Path $srcDir 'ThisDocument.cls'

if (-not $OutputPath) { $OutputPath = Join-Path $root 'ChatFormatter.dotm' }

foreach ($f in @($basFile, $clsFile)) {
    if (-not (Test-Path -LiteralPath $f)) { throw "Missing source: $f" }
}

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
        $bas = [System.IO.File]::ReadAllText($basFile, [System.Text.Encoding]::UTF8)
        # Drop the Attribute line, it is only valid on an imported .bas file.
        $bas = $bas -replace '(?m)^Attribute VB_Name.*\r?\n', ''

        $std = $vbProj.VBComponents.Add(1)   # vbext_ct_StdModule
        $std.Name = 'ChatFormatter'
        $std.CodeModule.AddFromString($bas)
        Write-Host "Injected module: ChatFormatter ($($std.CodeModule.CountOfLines) lines)"

        # ---- ThisDocument event code ----------------------------------------
        $cls = [System.IO.File]::ReadAllText($clsFile, [System.Text.Encoding]::UTF8)
        $cls = $cls -replace '(?s)^.*?Attribute VB_Customizable = True\r?\n', ''
        $cls = $cls -replace '(?m)^Attribute .*\r?\n', ''

        $thisDoc = $vbProj.VBComponents.Item('ThisDocument')
        $thisDoc.CodeModule.AddFromString($cls)
        Write-Host "Injected events: ThisDocument ($($thisDoc.CodeModule.CountOfLines) lines)"

        $doc.Save()
        Write-Host "Saved."

        # -----------------------------------------------------------------------
        # Compile check
        #
        # Word happily saves a template whose VBA does not compile; it only
        # complains later, as a modal "Compile error in hidden module" dialog
        # that blocks every macro run. Compile here so a broken build fails
        # loudly at build time instead.
        # -----------------------------------------------------------------------
        # When the project has errors, Word pops a modal dialog and Execute
        # throws E_FAIL. The dialog is still on screen at that point, so
        # read it in either case.
        try {
            [void]$word.VBE.CommandBars.FindControl(1, 578).Execute()   # Compile project
        }
        catch {
            Write-Host "Compile command returned: $($_.Exception.Message)"
        }

        $compileErrors = [WordDialogs]::CollectAndDismiss()

        if ($compileErrors) {
            Write-Host ''
            Write-Host 'COMPILE ERRORS:'
            $compileErrors | ForEach-Object { Write-Host "  $_" }
            throw 'ChatFormatter failed to compile; template not usable.'
        }
        Write-Host 'Compile: OK'
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

Write-Host ''
Write-Host "Build complete: $OutputPath"
Write-Host "Next: run install.bat to deploy it into Word's STARTUP folder."