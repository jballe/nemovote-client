New-Variable -Name NemoVoteContext -Value ([PSCustomObject]@{
    ServerUrl = $Null
    Token = $Null
}) -Scope Script -Force

$ErrorActionPreference = "STOP"

function Open-NemoVote {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=1 )]
        $ServerUrl,
        [Parameter(Mandatory=$true, Position=2 )]
        $Username,
        [Parameter(Mandatory=$true, Position=3 )]
        [SecureString] $Password,
        $Language = "da"
    )

    Set-NemoVoteServerUrl $ServerUrl

    $payload = [PSCustomObject]@{
        lang = $Language
        username = $Username
        password = $(ConvertFrom-SecureString $Password -AsPlainText)
    }

    $body = [System.Text.Encoding]::UTF8.GetBytes(($payload | ConvertTo-Json))
    $response = Invoke-RestMethod -Uri "${ServerUrl}/api/v1/auth/login" `
        -ContentType "application/json; charset=utf-8" `
        -Method POST `
        -Body $body
    
    HandleError -Response $response -Name "Login" -RequestObject $payload
    $token = $response.data.token | ConvertTo-SecureString -AsPlainText -Force
    Set-NemoVoteToken $token
}

function Get-NemoVoteServerUrl {
    $NemoVoteContext = (Get-Variable -Name NemoVoteContext -Scope Script).Value
    $NemoVoteContext.ServerUrl
}

function Set-NemoVoteServerUrl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=1 )]
        $ServerUrl
    )

    $NemoVoteContext = (Get-Variable -Name NemoVoteContext -Scope Script).Value
    $NemoVoteContext.ServerUrl = $ServerUrl.TrimEnd('/')
    Set-Variable -Name NemoVoteContext -Scope Script -Value $NemoVoteContext
}

function Set-NemoVoteToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=1 )]
        [SecureString]$Token
    )

    $NemoVoteContext = (Get-Variable -Name NemoVoteContext -Scope Script).Value
    $NemoVoteContext.Token = $Token
    Set-Variable -Name NemoVoteContext -Scope Script -Value $NemoVoteContext
}

function Get-NemoVoteToken {
    $NemoVoteContext = (Get-Variable -Name NemoVoteContext -Scope Script).Value
    $NemoVoteContext.Token
}
