param(
    $NemoVoteUrl = (get-content (Join-Path $PSSCriptRoot "../data.json" -Resolve) | convertFrom-json | Select-Object -ExpandProperty NemoVoteUrl),
    $NemoVoteUsername = (get-content (Join-Path $PSSCriptRoot "../data.json" -Resolve) | convertFrom-json | Select-Object -ExpandProperty NemoVoteUsername),
    [SecureString]$NemoVotePassword = (get-content (Join-Path $PSSCriptRoot "../data.json" -Resolve) | convertFrom-json | Select-Object -ExpandProperty NemoVotePassword | ConvertTo-SecureString -AsPlainText -Force)
)

$ErrorActionPreference = "STOP"

Import-Module (Join-Path $PSScriptRoot "../src/NemoVoteClient" -Resolve) -RequiredVersion 1.0.0 -Force

Open-NemoVote $NemoVoteUrl $NemoVoteUsername $NemoVotePassword
#Set-NemoVoteToken ("eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.etc.etc" | ConvertTo-SecureString -AsPlainText -Force)

#Get-NemoVoteUsers | Format-Table -Property username, accessLevel, displayName, id
#$users = Get-NemoVoteUsers | Where-Object { $_.username -like "user*" -or $_.username -like "voter*" }
#$users | ForEach-Object { Remove-NemoVoteUser $_.id }

$no = Get-Random -Minimum 2 -Maximum 100
$date = Get-Date
$newUser = Add-NemoVoteUser -Username "user${no}" -Displayname "User ${no} ${date}, Random gruppe" -Email "user${no}@balle-net.dk" -Pwd "123456"#(Get-RandomPassword -Length 8)
$users = Get-NemoVoteUsers | Where-Object { $_.username -like "user*" -or $_.username -like "voter*" }
#$users | ForEach-Object { Remove-NemoVoteUser $_.id }
$users | Format-Table -Property username, id

#$users | ForEach-Object { Remove-NemoVoteUser $_.id }

#$lists = Get-NemoVotingLists 
#$lists | Format-Table -Property name, id
#$list = $lists | ? { $_.name -eq "1. ekstra stemme"} | select -First 1
#Add-NemoVotingListMembers -ListId $list.id -UserId @($newUser.id)
#