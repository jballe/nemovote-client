function Get-NemoVotingLists {
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
    [CmdletBinding()]
    param(
        $List
    )

    $server = Get-NemoVoteServerUrl
    $token = Get-NemoVoteToken

    $response = Invoke-RestMethod "${server}/api/v1/voting-list/update" -Method PUT -Body ($List | ConvertTo-Json) -ContentType "application/json" -Authentication Bearer -Token $token
    HandleError $response -Name "Update list" -RequestObject $List
}

function Add-NemoVotingListMembers {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)]
        [string]$ListId,
        [Parameter(Mandatory=$True)]
        [array]$UserIds
    )

    Write-Verbose ("For list {0} add {1} users: {2}" -f $ListId, $UserIds.Length, ($UserIds -join ","))

    $lists = Get-NemoVotingLists
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
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)]
        [string]$ListId,
        [Parameter()]
        [array]$UserIds
    )

    Write-Verbose ("For list {0} add {1} users: {2}" -f $ListId, $UserIds.Length, ($UserIds -join ","))

    $lists = Get-NemoVotingLists
    $list = $lists | Where-Object { $_.id -eq $ListId }
    $list.users = $UserIds

    Update-NemoVotingList ([PSCustomObject]@{
        id = $list.id
        name = $list.name
        users = $list.users
        weight = $list.weight
    })
}
