param(
    $ServicePrincipal = (get-content "data.json" | convertFrom-json | Select-Object -ExpandProperty ServicePrincipal),
    $CertificatePath = (Join-Path $PSScriptRoot "*.p12" -Resolve),
    $FileId = (get-content "data.json" | convertFrom-json | Select-Object -ExpandProperty FileId),
    $NemoVoteUrl = (get-content "data.json" | convertFrom-json | Select-Object -ExpandProperty NemoVoteUrl),
    $NemoVoteUsername = (get-content "data.json" | convertFrom-json | Select-Object -ExpandProperty NemoVoteUsername),
    [SecureString]$NemoVotePassword = (get-content "data.json" | convertFrom-json | Select-Object -ExpandProperty NemoVotePassword | ConvertTo-SecureString -AsPlainText)
)

$ErrorActionPreference = "STOP"

Install-Module UMN-Google -Scope CurrentUser
Import-Module UMN-Google -Scope Local

# Set security protocol to TLS 1.2 to avoid TLS errors
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
# Google API Authozation
$scope = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/drive.file"
$certPswd = 'notasecret'
$accessToken = Get-GOAuthTokenService -scope $scope -iss $ServicePrincipal -certPath $CertificatePath -certPswd $certPswd
$properties = Get-GSheetSpreadSheetProperties -accessToken $accessToken -spreadSheetID $FileId
$sheetName = $properties.sheets.properties | Where-Object { $_.sheetId -eq 0 } | Select-Object -ExpandProperty title
$data = Get-GSheetData -accessToken $accessToken -spreadSheetID $FileId -sheetName $sheetName -rangeA1 "a6:e500" -cell "range"

$voters = $data | Where-object { $_.Navn -ne "" } | Group-Object -Property Email | ForEach-Object { [PSCustomObject]@{
    Name = ($_.Group | Select-Object -first 1).Navn
    Group = ($_.Group | Select-Object -first 1).Gruppenavn
    Email = ($_.Group | Select-Object -first 1).Email
    Votes = $_.Count
    UserId = $Null
} }
Write-Host ("Have read {0} voters" -f $voters.Count)

Import-Module (Join-Path $PSScriptRoot "./src/NemoVoteClient" -Resolve) -RequiredVersion 1.0.0 -Force

Open-NemoVote $NemoVoteUrl $NemoVoteUsername $NemoVotePassword

$existingUsers = Get-NemoVoteUsers
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

$additional1Vote = @() + ($mapped | Where-Object { $_.Votes -ge 2 } | Select-Object -ExpandProperty UserId)
$additional2Vote = @() + ($mapped | Where-Object { $_.Votes -ge 3 } | Select-Object -ExpandProperty UserId)
$additional3Vote = @() + ($mapped | Where-Object { $_.Votes -ge 4 } | Select-Object -ExpandProperty UserId)


$lists = Get-NemoVotingLists
$additional1ListId = $lists | Where-Object { $_.name -eq "1. ekstra stemme" } | Select-Object -ExpandProperty id
$additional2ListId = $lists | Where-Object { $_.name -eq "2. ekstra stemme" } | Select-Object -ExpandProperty id
$additional3ListId = $lists | Where-Object { $_.name -eq "3. ekstra stemme" } | Select-Object -ExpandProperty id

Set-NemoVotingListMembers -ListId $additional1ListId -UserIds $additional1Vote
Set-NemoVotingListMembers -ListId $additional2ListId -UserIds $additional2Vote
Set-NemoVotingListMembers -ListId $additional3ListId -UserIds $additional3Vote
