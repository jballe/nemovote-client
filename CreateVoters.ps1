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
$voters = $data | Where-object { $_.Navn -ne "" } | Group-Object -Property Email | ForEach-Object { [PSCustomObject]@{
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
        Write-Host "User ${email} will be created"
        $name = $_.Name
        $grp = $_.Group

        $user = Add-NemoVoteUser -Username $email -Email $email -DisplayName "${name}, ${grp}" -Pwd (Get-RandomPassword -Length 10)
        $existingUsers += $user
    } else {
        Write-Host "User ${email} is already created"
    }

    $_.UserId = $user.id
    $_
}

Write-Host "Remove existing users without a vote" -ForegroundColor Green
$emailsToDelete = Compare-Object -ReferenceObject $voters -DifferenceObject $existingUsers -Property email `
                | Where-Object { $_.SideIndicator -eq "=>" } `
                | Select-Object -ExpandProperty email
$existingUsers | Where-Object { $emailsToDelete -contains $_.email } | ForEach-Object {
    Remove-NemoVoteUser $_.id
}

Write-Host "Add users to voting lists" -ForegroundColor Green
# Figure out users who should be in additional voting lists
$additional1Vote = @() + ($mapped | Where-Object { $_.Votes -ge 2 } | Select-Object -ExpandProperty UserId)
$additional2Vote = @() + ($mapped | Where-Object { $_.Votes -ge 3 } | Select-Object -ExpandProperty UserId)
$additional3Vote = @() + ($mapped | Where-Object { $_.Votes -ge 4 } | Select-Object -ExpandProperty UserId)

# Add users to voting lists
$lists = Get-NemoVotingLists
$additional1ListId = $lists | Where-Object { $_.name -eq "1. ekstra stemme" } | Select-Object -ExpandProperty id
$additional2ListId = $lists | Where-Object { $_.name -eq "2. ekstra stemme" } | Select-Object -ExpandProperty id
$additional3ListId = $lists | Where-Object { $_.name -eq "3. ekstra stemme" } | Select-Object -ExpandProperty id

Set-NemoVotingListMembers -ListId $additional1ListId -UserIds $additional1Vote
Set-NemoVotingListMembers -ListId $additional2ListId -UserIds $additional2Vote
Set-NemoVotingListMembers -ListId $additional3ListId -UserIds $additional3Vote

Write-Host "Done!" -ForegroundColor Green
