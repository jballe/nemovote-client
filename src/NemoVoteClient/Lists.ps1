function Get-NemoVotingList {
    [CmdletBinding()]
    param(
    )

    $server = Get-NemoVoteServerUrl
    $token = Get-NemoVoteToken

    $response = Invoke-RestMethod "${server}/api/v1/voting-list/getall" -Authentication Bearer -Token $token
    HandleError $response

    $response.data
}

function Update-NemoVotingList {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        $List
    )

    $server = Get-NemoVoteServerUrl
    $token = Get-NemoVoteToken
    $body = [System.Text.Encoding]::UTF8.GetBytes( ($List | ConvertTo-Json) )

    $url = "${server}/api/v1/voting-list/update"
    If ($PSCmdlet.ShouldProcess($List, "Updating list")) {
        $response = Invoke-RestMethod $url -Method PUT -Body $body -ContentType "application/json; charset=utf-8" -Authentication Bearer -Token $token
        HandleError $response -Name "Update list" -RequestObject $List
    } else {
        Write-Verbose ("Will call $url with content: `n{0}" -f ($List | ConvertTo-Json))
    }
}

function Add-NemoVotingListMember {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)]
        [string]$ListId,
        [Parameter(Mandatory=$True)]
        [array]$UserIds
    )

    Write-Verbose ("For list {0} add {1} users: {2}" -f $ListId, $UserIds.Length, ($UserIds -join ","))

    $lists = Get-NemoVotingList
    $list = $lists | Where-Object { $_.id -eq $ListId }
    $list.users += $UserIds

    Update-NemoVotingList ([PSCustomObject]@{
        id = $list.id
        name = $list.name
        users = $list.users
        weight = $list.weight
    })
}

function Set-NemoVotingListMembers {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='It must update all members')]
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$True)]
        [string]$ListId,
        [Parameter()]
        [array]$UserIds
    )

    Write-Verbose ("For list {0} add {1} users: {2}" -f $ListId, $UserIds.Length, ($UserIds -join ","))

    $lists = Get-NemoVotingList
    $list = $lists | Where-Object { $_.id -eq $ListId }
    $list.users = $UserIds

    If ($PSCmdlet.ShouldProcess($ListId, "Updating list")) {
        Update-NemoVotingList ([PSCustomObject]@{
            id = $list.id
            name = $list.name
            users = $list.users
            weight = $list.weight
        })
    } else {
        Write-Verbose "Would call update list..."        
    }
}
