function Get-UcOneDriveWithMultiplePermissions {
    param(
        [string]$OutputPath,
        [switch]$MultiGeo
    )
    Write-Warning "Get-UcOneDriveWithMultiplePermissions will be deprecated in a future release, please use the Export-UcOneDriveWithMultiplePermissions instead."
    Export-UcOneDriveWithMultiplePermissions @params
}