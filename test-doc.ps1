$ErrorActionPreference = 'Stop'

$dotm = 'G:\Documents\Default Project\chat-formatter-word\ChatFormatter.dotm'
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
$sample = ($lines -join "")

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    # Open dotm as document
    $doc = $word.Documents.Open($dotm, $true, $false, $true)
    $doc.Content.Text = $sample

    Write-Host "BEFORE tables: $($doc.Tables.Count)"
    Write-Host ("BEFORE text:" )
    Write-Host $doc.Content.Text

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