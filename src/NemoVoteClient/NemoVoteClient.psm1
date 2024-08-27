Set-StrictMode -Version 3.0

$files = @(Get-ChildItem -Path $PSScriptRoot -Include "*.ps1" -File -Recurse)

($files) | ForEach-Object {
    try
    {
        Write-Verbose "Importing $_"
        . $_.FullName
    }
    catch
    {
        Write-Error $_.Exception.Message
    }
}
