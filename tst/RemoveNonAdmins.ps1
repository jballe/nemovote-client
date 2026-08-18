[Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingUserNameAndPassWordParams", '')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', 'Password', Justification = 'Obsolete')]
param(
    $NemoVoteUrl = (get-content (Join-Path $PSSCriptRoot "../data.json" -Resolve) | convertFrom-json | Select-Object -ExpandProperty NemoVoteUrl),
    $NemoVoteUsername = (get-content (Join-Path $PSSCriptRoot "../data.json" -Resolve) | convertFrom-json | Select-Object -ExpandProperty NemoVoteUsername),
    [string]$NemoVotePassword = (get-content (Join-Path $PSSCriptRoot "../data.json" -Resolve) | convertFrom-json | Select-Object -ExpandProperty NemoVotePassword)
)

$ErrorActionPreference = "STOP"

Import-Module (Join-Path $PSScriptRoot "../src/NemoVoteClient" -Resolve) -RequiredVersion 1.0.0 -Force

Open-NemoVote $NemoVoteUrl $NemoVoteUsername $NemoVotePassword

$users = @() + (Get-NemoVoteUser | Where-Object { $_.accessLevel -eq 1 })
Write-Host ("Will now remove {0} users:" -f $users.Length)
$users | Format-Table
Write-Host "Running..."
foreach($user in $users) {
    Write-Host "." -NoNewline
    Remove-NemoVoteUser -id $user.id
}

Write-Host ""
Write-Host -ForegroundColor Green "Done!"
