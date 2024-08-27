New-Variable -Name NemoVoteContext -Value ([PSCustomObject]@{
    ServerUrl = $Null
    Token = $Null
}) -Scope Script -Force

$ErrorActionPreference = "STOP"

function Open-NemoVote {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingConvertToSecureStringWithPlainText", '')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingUserNameAndPassWordParams", '')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', 'Password', Justification = 'Obsolete')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=1 )]
        [String] $ServerUrl,
        [Parameter(Mandatory=$true, Position=2 )]
        [String] $Username,
        [Parameter(Mandatory=$true, Position=3 )]
        [String] $Password,
        [String] $Language = "da"
    )

    Set-NemoVoteServerUrl $ServerUrl

    $payload = [PSCustomObject]@{
        lang = $Language
        username = $Username
        password = $Password
    }

    $url = ("{0}/api/v1/auth/login" -f (Get-NemoVoteServerUrl))
    $response = Invoke-RestMethod -Uri $url `
        -ContentType "application/json; charset=utf-8" `
        -Method POST `
        -Body ($payload | ConvertTo-Json)
    
    HandleError -Response $response -Name "Login" -RequestObject $payload
    $token = $response.data.token | ConvertTo-SecureString -AsPlainText -Force
    Set-NemoVoteToken $token
}

function Get-NemoVoteServerUrl {
    $NemoVoteContext = (Get-Variable -Name NemoVoteContext -Scope Script).Value
    $NemoVoteContext.ServerUrl
}

function Set-NemoVoteServerUrl {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='No side effects')]
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
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='No side effects')]
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
