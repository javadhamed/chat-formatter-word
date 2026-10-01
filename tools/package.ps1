<#
.SYNOPSIS
    Builds ChatFormatter.dotm and packages it into a release ZIP.

.DESCRIPTION
    Runs the build, then stages the files a user needs and zips them.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\package.ps1
#>

[CmdletBinding()]
param(
    [string]$Version = '1.0.0'
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot

Write-Host '==> Building template'
& (Join-Path $PSScriptRoot 'build.ps1')
if ($LASTEXITCODE -ne 0) { throw "build.ps1 failed" }

$stage = Join-Path $env:TEMP "cf-release-$Version"
if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force }
New-Item -ItemType Directory -Path $stage | Out-Null

Write-Host '==> Staging files'
foreach ($f in @('install.bat', 'uninstall.bat', 'ChatFormatter.dotm', 'README.md', 'LICENSE')) {
    Copy-Item -LiteralPath (Join-Path $root $f) -Destination $stage
}
Copy-Item -LiteralPath (Join-Path $root 'src') -Destination $stage -Recurse
Copy-Item -LiteralPath (Join-Path $root 'tools') -Destination $stage -Recurse

$zip = Join-Path $root "ChatFormatter-$Version.zip"
if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }

Write-Host '==> Zipping'
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip -Force

Remove-Item -LiteralPath $stage -Recurse -Force

$size = [math]::Round((Get-Item -LiteralPath $zip).Length / 1KB, 1)
Write-Host ''
Write-Host "Release ZIP: $zip ($size KB)"