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

function Get-NemoVoteAuthUrl {
    $NemoVoteContext = (Get-Variable -Name NemoVoteContext -Scope Script).Value
    $NemoVoteContext.AuthUrl
}

function Set-NemoVoteAuthUrl {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='No side effects')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=1 )]
        $AuthUrl
    )

    $NemoVoteContext = (Get-Variable -Name NemoVoteContext -Scope Script).Value
    $NemoVoteContext.AuthUrl = $AuthUrl.TrimEnd('/')
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
