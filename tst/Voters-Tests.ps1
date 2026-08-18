[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', 'NemoVotePassword', Justification = 'Obsolete')]
param(
    $NemoVoteUrl = (get-content (Join-Path $PSScriptRoot "../data.json" -Resolve) | convertFrom-json | Select-Object -ExpandProperty NemoVoteUrl),
    $NemoVoteUsername = (get-content (Join-Path $PSScriptRoot "../data.json" -Resolve) | convertFrom-json | Select-Object -ExpandProperty NemoVoteUsername),
    $NemoVotePassword = (get-content (Join-Path $PSScriptRoot "../data.json" -Resolve) | convertFrom-json | Select-Object -ExpandProperty NemoVotePassword)
)

$ErrorActionPreference = "STOP"

Import-Module (Join-Path $PSScriptRoot "../src/NemoVoteClient" -Resolve) -RequiredVersion 1.0.0 -Force

Write-Host "Now authenticating..."
Open-NemoVote $NemoVoteUrl $NemoVoteUsername $NemoVotePassword

#Get-NemoVoteUser | Format-Table -Property username, accessLevel, displayName, id
#$users = Get-NemoVoteUser | Where-Object { $_.username -like "user*" -or $_.username -like "voter*" }
#$users | ForEach-Object { Remove-NemoVoteUser $_.id }

$no = Get-Random -Minimum 2 -Maximum 100
$date = (Get-Date).Ticks
$username = "user${no}"
Write-Host "Now creating user $username"
$newUser = Add-NemoVoteUser -Username $username -Displayname "User ${no} ${date}, Random gruppe, æøåÆØÅ" -Email "nemovote-user-${no}@balle.dev"

Write-Host "Listing users"
$users = Get-NemoVoteUser -SearchQuery $username

#Update-NemoVoteUser -User $newUser
$users = Get-NemoVoteUser -SearchQuery $username

Write-Host "Get users"
#$users | ForEach-Object { Remove-NemoVoteUser $_.id }
$users | Format-Table -Property username, id, displayname

Write-Host "Now removing user"
Remove-NemoVoteUser -Id $newUser.id

#$users | ForEach-Object { Remove-NemoVoteUser $_.id }

#$lists = Get-NemoVotingList
#$lists | Format-Table -Property name, id
#$list = $lists | ? { $_.name -eq "1. ekstra stemme"} | select -First 1
#Add-NemoVotingListMember -ListId $list.id -UserId @($newUser.id)
#