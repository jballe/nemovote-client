[CmdletBinding]
function HandleError {
    param(
        [Parameter(Mandatory=$true, Position=1)]
        $Response,
        [Parameter(Mandatory=$false, Position=2)]
        $Name = "method",
        [Parameter(Mandatory=$false, Position=3)]
        $RequestObject = $Null

    )

    If($Null -ne $RequestObject) {
        Write-Verbose "Sent request to $Name"
        $RequestObject | ConvertTo-Json | Write-Verbose
    }
    If ($Response.code -ne 0) {
        Write-Warning "Result from $Name"
        $response | ConvertTo-Json | Write-Warning
    }
}

