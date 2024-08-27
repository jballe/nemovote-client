param(
    $PackageName = "NemoVoteClient"
)
$ErrorActionPreference = "STOP"

Uninstall-Module -Name $PackageName -ErrorAction SilentlyContinue
Write-Host "Done" -ForegroundColor Green
