param(
    $ServicePrincipal = (get-content "data.json" | convertFrom-json | Select-Object -ExpandProperty ServicePrincipal),
    $CertificatePath = (Join-Path $PSScriptRoot "*.p12" -Resolve),
    $FileId = (get-content "data.json" | convertFrom-json | Select-Object -ExpandProperty FileId),
    $NemoVoteUrl = (get-content "data.json" | convertFrom-json | Select-Object -ExpandProperty NemoVoteUrl),
    $NemoVoteUsername = (get-content "data.json" | convertFrom-json | Select-Object -ExpandProperty NemoVoteUsername),
    [SecureString]$NemoVotePassword = (get-content "data.json" | convertFrom-json | Select-Object -ExpandProperty NemoVotePassword | ConvertTo-SecureString -AsPlainText -Force)
)

$ErrorActionPreference = "STOP"

Write-Host "Powershell version:" -ForegroundColor Green
(Get-Host).Version | Format-Table

# Set security protocol to TLS 1.2 to avoid TLS errors
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

Write-Host "Install Powershell module" -ForegroundColor Green
Install-Module UMN-Google -Scope CurrentUser -Force
Import-Module UMN-Google -Scope Local -Force

Write-Host "Google API Authentication" -ForegroundColor Green
$scope = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/drive.file"
$certPswd = 'notasecret'
$accessToken = Get-GOAuthTokenService -scope $scope -iss $ServicePrincipal -certPath $CertificatePath -certPswd $certPswd

Write-Host "Read data from Google sheet" -ForegroundColor Green
$properties = Get-GSheetSpreadSheetProperties -accessToken $accessToken -spreadSheetID $FileId
$sheetName = $properties.sheets.properties | Where-Object { $_.sheetId -eq 0 } | Select-Object -ExpandProperty title
$data = Get-GSheetData -accessToken $accessToken -spreadSheetID $FileId -sheetName $sheetName -rangeA1 "a6:e500" -cell "range"

# Process data
$voters = $data | Where-object { $_.Navn -ne "" -and $_.Gruppenavn -ne "" } | Group-Object -Property Email | ForEach-Object { [PSCustomObject]@{
    Name = ($_.Group | Select-Object -first 1).Navn
    Group = ($_.Group | Select-Object -first 1).Gruppenavn
    Email = ($_.Group | Select-Object -first 1).Email
    Votes = $_.Count
    UserId = $Null
} }
Write-Host ("Have read {0} voters" -f $voters.Count)

Write-Host "Login to NemoVote" -ForegroundColor Green
Import-Module (Join-Path $PSScriptRoot "./src/NemoVoteClient" -Resolve) -RequiredVersion 1.0.0 -Force
Open-NemoVote $NemoVoteUrl $NemoVoteUsername $NemoVotePassword

# Create missing users and get userId of all users
Write-Host "Ensure users are created" -ForegroundColor Green
$existingUsers = Get-NemoVoteUsers | Where-Object { $_.accessLevel -eq 1 }
$mapped = $voters | ForEach-Object {
    $email = $_.Email
    $user = $existingUsers | Where-Object { $_.email -eq $email }
    if($user -eq $Null) {
        $name = $_.Name
        $grp = $_.Group
        $displayName = "${name}, ${grp}"
        Write-Host "User ${displayName} will be created"

        $user = Add-NemoVoteUser -Username $email -Email $email -DisplayName $displayName -Pwd (Get-RandomPassword -Length 10)
        $existingUsers += $user
        Send-NemoVoteUserCredentials $user.id
    } else {
        #Write-Host "User ${email} is already created"
    }

    $_.UserId = $user.id
    $_
}

Write-Host "Remove existing users without a vote" -ForegroundColor Green
$emailsToDelete = Compare-Object -ReferenceObject $voters -DifferenceObject $existingUsers -Property email `
                | Where-Object { $_.SideIndicator -eq "=>" } `
                | Select-Object -ExpandProperty email
$usersToDelete = $existingUsers | Where-Object { $emailsToDelete -contains $_.email }
$usersToDelete | Format-Table
$usersToDelete | ForEach-Object {
    Remove-NemoVoteUser $_.id
}

Write-Host "Add users to voting lists" -ForegroundColor Green
# Figure out users who should be in additional voting lists
$additional1Vote = @() + ($mapped | Where-Object { $_.Votes -ge 2 } | Select-Object -ExpandProperty UserId)
$additional2Vote = @() + ($mapped | Where-Object { $_.Votes -ge 3 } | Select-Object -ExpandProperty UserId)
$additional3Vote = @() + ($mapped | Where-Object { $_.Votes -ge 4 } | Select-Object -ExpandProperty UserId)

# Add users to voting lists
$lists = Get-NemoVotingLists
$additional1List = $lists | Where-Object { $_.name -eq "1. ekstra stemme" }
$additional2List = $lists | Where-Object { $_.name -eq "2. ekstra stemme" }
$additional3List = $lists | Where-Object { $_.name -eq "3. ekstra stemme" }

Write-Host ("Set {0} users for voting list {1}" -f $additional1Vote.Count, $additional1List.name)
Set-NemoVotingListMembers -ListId $additional1List.id -UserIds $additional1Vote
Write-Host ("Set {0} users for voting list {1}" -f $additional2Vote.Count, $additional2List.name)
Set-NemoVotingListMembers -ListId $additional2List.id -UserIds $additional2Vote
Write-Host ("Set {0} users for voting list {1}" -f $additional3Vote.Count, $additional3List.name)
Set-NemoVotingListMembers -ListId $additional3List.id -UserIds $additional3Vote

Write-Host "Done!" -ForegroundColor Green
