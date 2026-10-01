$ErrorActionPreference = 'Stop'
$dotmPath = 'G:\Documents\Default Project\chat-formatter-word\ChatFormatter.dotm'

$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$d = $w.Documents.Add()
$d.Content.Text = '# H
**bold**
| A | B |
|---|---|
| 1 | 2 |
`code`'

for ($n = 1; $n -le 4; $n++) {
    Write-Host ('=== Stage ' + $n + ' ===')
    try {
        $w.Run('ChatFormatter.CF_TestStage', [ref]$n)
        Write-Host 'OK'
    }
    catch {
        Write-Host ('ERR: ' + $_.Exception.Message)
    }
}

$d.Close([ref]0)
$w.Quit()