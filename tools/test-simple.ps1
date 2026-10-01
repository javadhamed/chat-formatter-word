$ErrorActionPreference = 'Stop'
$dotmPath = 'G:\Documents\Default Project\chat-formatter-word\ChatFormatter.dotm'

for ($n = 1; $n -le 4; $n++) {
    Write-Host ('=== Stage ' + $n + ' ===')
    $job = Start-Job -ArgumentList $dotmPath, $n -ScriptBlock {
        param($d, $n)
        $w = New-Object -ComObject Word.Application
        $w.Visible = $false
        $w.DisplayAlerts = 0
        $doc = $w.Documents.Add()
        $doc.Content.Text = '# H
**bold**
| A | B |
|---|---|
| 1 | 2 |
`code`'
        $w.Run('ChatFormatter.CF_TestStage', [ref]$n)
        $doc.Close([ref]0)
        $w.Quit()
    }
    $ok = Wait-Job $job -Timeout 20
    if ($ok) {
        Receive-Job $job | Out-Null
        Write-Host ('Stage ' + $n + ': OK')
    } else {
        Write-Host ('Stage ' + $n + ': TIMEOUT')
        Stop-Job $job
    }
    Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2
}