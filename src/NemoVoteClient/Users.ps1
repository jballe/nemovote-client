function Get-NemoVoteUser {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseOutputTypeCorrectly', '', Justification='Transparent type')]
    [CmdletBinding()]
    param(
        $SearchQuery,
        $PageSize = 50
    )

    $server = Get-NemoVoteServerUrl
    $token = Get-NemoVoteToken

    $result = @()
    $page = 0
    $totalUsers = 0
    do {
        $page++
        $url = "${server}/api/v1/user/getall?page=${page}&pageSize=${PageSize}&searchQuery=${SearchQuery}"
        $response = Invoke-RestMethod $url -Authentication Bearer -Token $token
        HandleError $response

        $totalUsers = $response.data.totalLength
        $userPage = $response.data.collection
        $result += @() + $userPage
    } while($totalUsers -gt $result.Length -and $userPage.Length -gt 0)

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

    $body = [System.Text.Encoding]::UTF8.GetBytes( ($payload | ConvertTo-Json) )
    $response = Invoke-RestMethod "${server}/api/v1/user/create" -Method POST -Body $body -ContentType "application/json; charset=utf-8" -Authentication Bearer -Token $token
    HandleError -Response $response -Name "Add-NemoVoteUser" -RequestObj $payload
    if("data" -in $response.PSObject.Properties.Name) {
        $response.data
    }
}

function Update-NemoVoteUser {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true, Position=1, ValueFromPipeline)]
        $User
    )
    
    $json = ($User | ConvertTo-Json)
    $body = [System.Text.Encoding]::UTF8.GetBytes( $json )
    $server = Get-NemoVoteServerUrl
    $token = Get-NemoVoteToken

    $url = "${server}/api/v1/user/update"
    If ($PSCmdlet.ShouldProcess($User, "Updating user")) {
        $response = Invoke-RestMethod $url -Method PUT -Body $body -ContentType "application/json; charset=utf-8" -Authentication Bearer -Token $token
        HandleError $response -Name "Update user" -RequestObject $User
    } else {
        Write-Verbose "Skip calling $url with content $json"
    }
}

function Remove-NemoVoteUser {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true, Position=1, ValueFromPipelineByPropertyName)]
        $Id
    )

    $server = Get-NemoVoteServerUrl
    $token = Get-NemoVoteToken

    Write-Verbose "Delete user ${Id}"
    $url = "${server}/api/v1/user/delete/${Id}"
    If ($PSCmdlet.ShouldProcess($Id, "Deleting user")) {
        $response = Invoke-RestMethod  -Method DELETE -Authentication Bearer -Token $token
        HandleError $response
    } else {
        Write-Verbose "Skip calling delete on $url"
    }
}

function Send-NemoVoteUserCredentials {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Correct naming')]
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=1)]
        $Id
    )

    $server = Get-NemoVoteServerUrl
    $token = Get-NemoVoteToken

    $url = "${server}/api/v1/user/credentials/${Id}"
    $response = Invoke-RestMethod $url -Method POST -Authentication Bearer -Token $token
    HandleError $response
}