$ErrorActionPreference = 'Stop'
$dotmPath = 'G:\Documents\Default Project\chat-formatter-word\ChatFormatter.dotm'

$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0

$d = $w.Documents.Open($dotmPath, $true, $false, $true)
Write-Host 'Template opened as document'

$testText = '# H
**bold**
| A | B |
|---|---|
| 1 | 2 |
`code`'
$d.Content.Text = $testText

Write-Host ('Paragraphs: ' + $d.Paragraphs.Count)
for ($i = 1; $i -le $d.Paragraphs.Count; $i++) {
    $txt = $d.Paragraphs.Item($i).Range.Text
    Write-Host ('P' + $i + ': [' + $txt.Trim() + ']')
}

$d.Close([ref]0)
$w.Quit()