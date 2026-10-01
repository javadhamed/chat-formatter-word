$sample = "# Heading One`rSome **bold** and *italic* and ~~strike~~ and `` code `` here.`r`r## Heading Two`r`r> a quoted line`r`r- first bullet`r- second bullet`r`r1. one`r2. two`r`r| Name | Age |`r|------|-----|`r| Ali  | 30  |`r| Sara | 25  |`r`r``` `rplain code line`r```"
$dotm="G:\Documents\Default Project\chat-formatter-word\ChatFormatter.dotm"
$w=New-Object -ComObject Word.Application
$w.Visible=$false; $w.DisplayAlerts=0
$d=$w.Documents.Open($dotm, $true, $false, $true)
$d.Content.Text = $sample
Write-Host "BEFORE tables: $($d.Tables.Count)"
try{$w.Run('ChatFormatter.CF_TestFormatRange'); "OK"} catch{"ERR: " + $_.Exception.Message}
$err = $w.Run('ChatFormatter.CF_GetLastError')
"LastError: $err"
$d.Close([ref]0); $w.Quit()