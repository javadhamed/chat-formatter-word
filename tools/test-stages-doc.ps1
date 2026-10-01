$ErrorActionPreference = 'Stop'
$dotmPath = "G:\Documents\Default Project\chat-formatter-word\ChatFormatter.dotm"

$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0

$d = $w.Documents.Open($dotmPath, $true, $false, $true)
Write-Host "Template opened as document"

Write-Host "=== Stage 1 ==="
try {
    $w.Run('ChatFormatter.CF_TestStage', [ref]1)
    Write-Host "Stage 1: OK"
}
catch {
    Write-Host ("ERR: " + $_.Exception.Message)
}

Write-Host "=== Stage 2 ==="
try {
    $w.Run('ChatFormatter.CF_TestStage', [ref]2)
    Write-Host "Stage 2: OK"
}
catch {
    Write-Host ("ERR: " + $_.Exception.Message)
}

Write-Host "=== Stage 3 ==="
try {
    $w.Run('ChatFormatter.CF_TestStage', [ref]3)
    Write-Host "Stage 3: OK"
}
catch {
    Write-Host ("ERR: " + $_.Exception.Message)
}

Write-Host "=== Stage 4 ==="
try {
    $w.Run('ChatFormatter.CF_TestStage', [ref]4)
    Write-Host "Stage 4: OK"
}
catch {
    Write-Host ("ERR: " + $_.Exception.Message)
}

$d.Close([ref]0)
$w.Quit()