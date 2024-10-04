function Get-UcM365LicenseAssignment {
    param(
        [string]$SKU,    
        [switch]$UseFriendlyNames,
        [switch]$SkipServicePlan,
        [string]$OutputPath,
        [switch]$DuplicateServicePlansOnly
    )
    Write-Warning "Get-UcM365LicenseAssignment will be deprecated in a future release, please use the Export-UcM365LicenseAssignment instead."
    Export-UcM365LicenseAssignment @params
}