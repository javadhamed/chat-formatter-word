# Test each CF_TestStage individually
$ErrorActionPreference = 'Stop'

$dotmPath = "G:\Documents\Default Project\chat-formatter-word\ChatFormatter.dotm"
$stageNames = @{1='PipeTables'; 2='TSVTables'; 3='HorizontalRules'; 4='CodeBlocks'}

for ($n = 1; $n -le 4; $n++) {
    $stageName = $stageNames[$n]
    Write-Host "=== Testing Stage $n ($stageName) ==="

    $job = Start-Job -ArgumentList $dotmPath, $n -ScriptBlock {
        param($dotm, $n)
        $w = New-Object -ComObject Word.Application
        $w.Visible = $false
        $w.DisplayAlerts = 0
        $d = $w.Documents.Add()
        $d.Content.Text = "# H`r`n**bold**`r`n| A | B |`r`n|---|---|`r`n| 1 | 2 |`r`n```code```"
        $w.Run('ChatFormatter.CF_TestStage', [ref]$n)
        $d.Close([ref]0)
        $w.Quit()
    }

    $finished = Wait-Job $job -Timeout 20
    if ($finished) {
        Receive-Job $job | Out-Null
        Write-Host "Stage $n ($stageName): OK"
    } else {
        Write-Host "Stage $n ($stageName): TIMEOUT"
        Stop-Job $job
    }
    Get-Job | Remove-Job -Force -ErrorAction SilentlyContinue
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2
}