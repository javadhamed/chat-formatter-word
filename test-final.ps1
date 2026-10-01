# End-to-end test: opens ChatFormatter.dotm as a document and runs all stages
$ErrorActionPreference = 'Stop'

$dotm = Join-Path $PSScriptRoot 'ChatFormatter.dotm'
if (-not (Test-Path -LiteralPath $dotm)) { throw "Build first: $dotm not found" }

$bt = [string][char]96
$fence = $bt + $bt + $bt
$tick2 = $bt + $bt

$sample = @(
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
) -join "`r"

$names = @{
    1 = 'PipeTables'; 2 = 'TSVTables'; 3 = 'HorizontalRules'; 4 = 'CodeBlocks'
    5 = 'Headings'; 6 = 'Bold'; 7 = 'Italic'; 8 = 'Strike'; 9 = 'InlineCode'
    10 = 'Links'; 11 = 'Blockquotes'; 12 = 'Bullets'; 13 = 'Numbered'
    14 = 'StrayMarkers'
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    # Open dotm AS DOCUMENT (not add-in) - this works!
    $doc = $word.Documents.Open($dotm, $true, $false, $true)
    $doc.Content.Text = $sample

    Write-Host "Paragraphs: $($doc.Paragraphs.Count)"

    foreach ($n in 1..14) {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        Write-Host "BEGIN $($names[$n])"
        try {
            $word.Run('ChatFormatter.CF_TestStage', [ref]$n)
            Write-Host "OK $($names[$n]) $($sw.ElapsedMilliseconds)ms"
        }
        catch {
            Write-Host "THREW $($names[$n]): $($_.Exception.Message)"
        }
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