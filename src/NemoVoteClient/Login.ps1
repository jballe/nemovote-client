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
        [String] $Password
    )

    Set-NemoVoteServerUrl $ServerUrl

    $result = Open-NemoVoteWithToken -Username $Username -Password $Password -ServerUrl $ServerUrl
    $result | Out-Null
}
