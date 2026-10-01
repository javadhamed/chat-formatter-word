# Automated end-to-end test for ChatFormatter.dotm
# Creates a doc with AI-chat markdown, runs the macro, dumps the result.

$ErrorActionPreference = 'Stop'

$root = Join-Path (Split-Path -Parent $PSScriptRoot) ''
$dotm = Join-Path (Split-Path -Parent $PSScriptRoot) 'ChatFormatter.dotm'
if (-not (Test-Path -LiteralPath $dotm)) { throw "Build first: $dotm not found" }

$bt = [string][char]96
$fence = $bt + $bt + $bt
$tick2 = $bt + $bt
$tick1 = $bt

$lines = @(
    '# Heading One'
    "Some **bold** and *italic* and ~~strike~~ and $tick2 code $tick2 here."
    ''
    '## Heading Two'
    ''
    '> a quoted line'
    ''
    '- first bullet'
    '- second bullet'
    ''
    '1. one'
    '2. two'
    ''
    '| Name | Age |'
    '|------|-----|'
    '| Ali  | 30  |'
    '| Sara | 25  |'
    ''
    $fence
    'plain code line'
    $fence
)
# Word wants CR-only paragraph marks. A joined CRLF string arrives as one
# long line with the LFs swallowed, so normalise before assigning.
$sample = ($lines -join "`r")

# Install to STARTUP so the add-in loads. Any running Word holds the file
# open, which makes the copy silently fail and leaves the previous build
# loaded instead, so make sure nothing is holding it.
Get-Process WINWORD -ErrorAction SilentlyContinue |
    Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 800

$startup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
New-Item -ItemType Directory -Path $startup -Force | Out-Null
Copy-Item -LiteralPath $dotm -Destination $startup -Force

$sec = 'HKCU:\Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\LocationChatFormatter'
if (-not (Test-Path -LiteralPath $sec)) { New-Item -Path $sec -Force | Out-Null }
Set-ItemProperty -Path $sec -Name Path -Value $startup
Set-ItemProperty -Path $sec -Name AllowSubfolders -Value 1 -Type DWord

if (-not (Test-Path -LiteralPath 'HKCU:\Software\ChatFormatter')) {
    New-Item -Path 'HKCU:\Software\ChatFormatter' -Force | Out-Null
}
Set-ItemProperty -Path 'HKCU:\Software\ChatFormatter' -Name AutoFormat -Value '0'
Set-ItemProperty -Path 'HKCU:\Software\ChatFormatter' -Name Tables -Value '1'

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Add()
    $doc.Content.Text = $sample

    Write-Host "BEFORE tables: $($doc.Tables.Count)"
    Write-Host ("BEFORE text:" )
    Write-Host $doc.Content.Text

    # Run the silent test entry point from the loaded add-in
    $word.Run('ChatFormatter.CF_TestFormatRange')

    Start-Sleep -Seconds 2

    $err = $word.Run('ChatFormatter.CF_GetLastError')
    if ($err) {
        $parts = $err -split '\|', 2
        Write-Host "MACRO ERROR $($parts[0]): $($parts[1])"
    } else {
        Write-Host 'MACRO ERROR none'
    }

    Write-Host ''
    Write-Host "================= AFTER ================="
    Write-Host "tables: $($doc.Tables.Count)"
    Write-Host '--- text ---'
    Write-Host $doc.Content.Text

    if ($doc.Tables.Count -gt 0) {
        $t = $doc.Tables.Item(1)
        Write-Host "table rows=$($t.Rows.Count) cols=$($t.Columns.Count)"
        Write-Host "row1 = $($t.Rows.Item(1).Range.Text -replace [char]7,'|')"
    }

    Write-Host '--- styles used ---'
    $styles = @{}
    foreach ($p in $doc.Paragraphs) {
        $s = $p.Style.NameLocal
        if (-not $styles.ContainsKey($s)) { $styles[$s] = 0 }
        $styles[$s]++
    }
    $styles.GetEnumerator() | ForEach-Object { "  $($_.Key): $($_.Value)" }
}
finally {
    $doc.Close([ref]0)
    $word.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}