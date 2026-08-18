function Get-NemoVotingPoll {
    [CmdletBinding()]
    param(
    )

    $server = Get-NemoVoteServerUrl
    $token = Get-NemoVoteToken

    $response = Invoke-RestMethod "${server}/api/v1.1/vote/getall" -ContentType "application/json; charset=utf-8" -Authentication Bearer -Token $token
    HandleError $response -Name "Get-NemoVotingPoll"

    $response.data
}

function Get-NemoVotingPollResult {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 1, ValueFromPipelineByPropertyName)]
        $id
    )

    begin {
        $server = Get-NemoVoteServerUrl
    }

    process {
        $token = Get-NemoVoteToken

        $response = Invoke-RestMethod "${server}/api/v1.1/vote/result/${id}" -ContentType "application/json; charset=utf-8" -Authentication Bearer -Token $token
        HandleError $response -Name "Get-NemoVotingPollResult"

        $response.data
    }
}

function New-NemoVotingPoll {
    [CmdletBinding(SupportsShouldProcess = $True)]
    param(
        [Parameter(Mandatory)]
        $Name,
        $Description,

        [Parameter(Mandatory)]
        [Array]$VotingLists,

        [Array]$VoteChoices = @("votes.choices.inFavour", "votes.choices.against", "votes.choices.abstain"),
        $BallotPrivacy = "SECRET",
        $ResultPrivacy = "SECRET",
        $VotePrivacy = "SECRET",
        $WhoHasVotedStatus = "SECRET",
        $Status = "draft",
        [Switch]$SyncVotingRightsAfterOpened,
        $WeightCalculationType = "LIST",
        $AmountOfVoteChoicesMin = 1,
        $AmountOfVoteChoicesMax = 1,
        [Switch]$AutoClose
    )

    $payload = @{
        name                        = $Name
        description                 = $Description
        amountOfVoteChoices         = @{
            min = $AmountOfVoteChoicesMin
            max = $AmountOfVoteChoicesMax
        }
        autoClose                   = ($AutoClose -eq $True)
        ballotPrivacy               = $BallotPrivacy
        resultPrivacy               = $ResultPrivacy
        votePrivacy                 = $VotePrivacy
        whoHasVotedStatus          = $WhoHasVotedStatus

        voteChoices                 = $VoteChoices
        votingLists                 = $VotingLists

        weightCalculationType       = $WeightCalculationType
        syncVotingRightsAfterOpened = ($SyncVotingRightsAfterOpened -eq $True)
        type                        = "Poll"
        status                      = $Status
    }

    $server = Get-NemoVoteServerUrl
    $token = Get-NemoVoteToken

    $body = [System.Text.Encoding]::UTF8.GetBytes( ($payload | ConvertTo-Json) )

    $url = "${server}/api/v1.1/vote/create"
    If ($PSCmdlet.ShouldProcess("Updating list")) {
        $response = Invoke-RestMethod $url -Method POST -Body $body -ContentType "application/json" -Authentication Bearer -Token $token
        HandleError $response -RequestObject $payload -Name "New-NemoVotingPoll"
    }
    else {
        Write-Verbose ("Will call $url with content: `n{0}" -f ($List | ConvertTo-Json))
    }
}