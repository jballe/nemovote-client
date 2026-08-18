function Open-NemoVoteWithToken {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingConvertToSecureStringWithPlainText", '')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingUserNameAndPassWordParams", '')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', 'Password', Justification = 'Obsolete')]
    param (
        [Parameter(Mandatory=$true, Position=1 )]
        [string]$Username,
        [Parameter(Mandatory=$true, Position=2 )]
        [string]$Password,
        [string]$ServerUrl = "https://kfumspejderne.nemovote.com",
        [string]$ClientId = "nemovote-client",
        [string]$ClientSecret
    )

    try {
        # First, get the init data to find the authentication endpoint
        $initResponse = Invoke-RestMethod -Uri "$ServerUrl/api/init" -Method Get
        $authBaseUrl = $initResponse.data.authUrl
        $tokenEndpoint = "$authBaseUrl/protocol/openid-connect/token"
        
        # Prepare the authentication payload for KeyCloak
        $authBody = @{
            client_id = $ClientId
            username = $Username
            password = $Password
            grant_type = "password"
            scope = "openid"
        }
        
        if ($ClientSecret) {
            $authBody.client_secret = $ClientSecret
        }

        # Set headers for the authentication request
        $authHeaders = @{
            "Content-Type" = "application/x-www-form-urlencoded"
            "Accept" = "application/json"
        }

        # Convert body to form-urlencoded format
        $formBody = ($authBody.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "&"

        # Perform the authentication request
        $authResponse = Invoke-RestMethod -Uri $tokenEndpoint -Method Post -Body $formBody -Headers $authHeaders
        
        # Extract the access token from the response
        $token = $authResponse.access_token
        
        if (-not $token) {
            throw "Failed to obtain access token. Response: $($authResponse | ConvertTo-Json)"
        }

        $result = ($token | ConvertTo-SecureString -AsPlainText -Force)
        Set-NemoVoteToken $result
        return $result
    }
    catch {
        Write-Error "Error during authentication: $_"
        throw
    }
}