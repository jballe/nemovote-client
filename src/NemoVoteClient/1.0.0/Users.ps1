function Get-NemoVoteUsers {
    [CmdletBinding()]
    param(
    )

    $server = Get-NemoVoteServerUrl
    $token = Get-NemoVoteToken

    $pageSize = 50
    $result = @()
    $page = 0
    $totalUsers = 0
    do {
        $page++
        $response = Invoke-RestMethod "${server}/api/v1/user/getall?page=${page}&pageSize=${pageSize}" -Authentication Bearer -Token $token
        HandleError $response

        $totalUsers = $response.data.totalLength
        $userPage = $response.data.collection
        $userPage.GetType()
        $result += @() + $userPage
    } while($totalUsers -gt $result.Length -and $userPage.Length > 0)

    $result
}

function Add-NemoVoteUser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=1)]
        $Username,
        [Parameter(Mandatory=$true, Position=2)]
        $DisplayName,
        [Parameter(Mandatory=$false, Position=3)]
        $Email,
        [Parameter(Mandatory=$true, Position=4)]
        $Pwd = $Null,
        [Parameter(Mandatory=$false, Position=5)]
        [int]$AccessLevel=1,
        [Parameter(Mandatory=$false, Position=6)]
        [bool]$ForcePasswordChange = $True
    )

    $payload = @{
        accessLevel = $AccessLevel
        username = $Username
        email = $Email
        displayname = $DisplayName
        password = $Pwd
        forcePasswordChange = $ForcePasswordChange
    }

    $server = Get-NemoVoteServerUrl
    $token = Get-NemoVoteToken
    $response = Invoke-RestMethod "${server}/api/v1/user/create" -Method POST -Body $payload -Authentication Bearer -Token $token
    HandleError -Response $response -Name "Add-NemoVoteUser" -RequestObj $payload
    if("data" -in $response.PSObject.Properties.Name) {
        $response.data
    }
}

function Remove-NemoVoteUser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=1)]
        $Id
    )

    $server = Get-NemoVoteServerUrl
    $token = Get-NemoVoteToken

    Write-Verbose "Delete user ${Id}"
    $response = Invoke-RestMethod "${server}/api/v1/user/delete/${Id}" -Method DELETE -Authentication Bearer -Token $token
    HandleError $response
}