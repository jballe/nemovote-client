[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', 'NemoVotePassword', Justification = 'Obsolete')]
param(
    $NemoVoteUrl = (get-content (Join-Path $PSSCriptRoot "../data.json" -Resolve) | convertFrom-json | Select-Object -ExpandProperty NemoVoteUrl),
    $NemoVoteUsername = (get-content (Join-Path $PSSCriptRoot "../data.json" -Resolve) | convertFrom-json | Select-Object -ExpandProperty NemoVoteUsername),
    [string]$NemoVotePassword = (get-content (Join-Path $PSSCriptRoot "../data.json" -Resolve) | convertFrom-json | Select-Object -ExpandProperty NemoVotePassword)
)

$ErrorActionPreference = "STOP"

Import-Module (Join-Path $PSScriptRoot "../src/NemoVoteClient" -Resolve) -RequiredVersion 1.0.0 -Force

Open-NemoVote -ServerUrl $NemoVoteUrl -Username $NemoVoteUsername -Password $NemoVotePassword

Write-Host "Received securetoken"