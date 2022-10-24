<#
.SYNOPSIS
Validate which Intune Compliance policies are supported by Microsoft Teams Android Devices

.DESCRIPTION
This function will validate each setting in the Intune Compliance Policy to make sure they are in line with the supported settings:

    https://docs.microsoft.com/en-us/microsoftteams/rooms/supported-ca-and-compliance-policies?tabs=phones#supported-device-compliance-policies

Contributors: Traci Herr, David Paulino

Requirements: Microsoft Graph PowerShell Module (Install-Module Microsoft.Graph)

.PARAMETER Detailed
Displays test results for all settings in each Intune Compliance Policy

.EXAMPLE 
PS> Test-UcTeamsDevicesCompliancePolicy

.EXAMPLE 
PS> Test-UcTeamsDevicesCompliancePolicy -Detailed

#>
Function Test-UcTeamsDevicesCompliancePolicy {

    Param(
        [Parameter(Mandatory = $false)]
        [switch]$Detailed,
        [Parameter(Mandatory = $false)]
        [string]$PolicyID,
        [Parameter(Mandatory = $false)]
        [string]$PolicyName
    )

    $connectedMSGraph = $false
    $CompliancePolicies = $null

    $scopes = (Get-MgContext).Scopes

    if (!($scopes) -or !( "DeviceManagementConfiguration.Read.All" -in $scopes )) {
        Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All"
    }

    try {
        $CompliancePolicies = (Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies" -Method GET).value
        $connectedMSGraph = $true
    }
    catch {
        Write-Error 'Please connect to MS Graph with Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All" before running this script'
    }

    if ($connectedMSGraph) {
        $output = [System.Collections.ArrayList]::new()
        $outputSum = [System.Collections.ArrayList]::new()
        foreach ($CompliancePolicy in $CompliancePolicies) {
            if ((($PolicyID -eq $CompliancePolicy.id) -or ($PolicyName -eq $CompliancePolicy.displayName) -or (!$PolicyID -and !$PolicyName)) -and ($CompliancePolicy."@odata.type" -eq "#microsoft.graph.androidCompliancePolicy")) {
                $PolicyErrors = 0
                $PolicyWarnings = 0

                $ID = 1
                $Setting = "deviceThreatProtectionEnabled"
                $Comment = ""
                if ($CompliancePolicy.deviceThreatProtectionEnabled) {
                    $Status = "Unsupported"
                    $PolicyErrors++
                }
                else {
                    $Status = "Supported"
                }

                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    ID                 = $ID
                    PolicyID           = $CompliancePolicy.id
                    PolicyName         = $CompliancePolicy.displayName
                    Setting            = $Setting
                    Value              = $CompliancePolicy.deviceThreatProtectionEnabled
                    TeamsDevicesStatus = $Status 
                    Comment            = $Comment
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceCompliancePolicy')
                $output.Add($SettingPSObj) | Out-Null


                $ID = 3
                $Setting = "securityBlockJailbrokenDevices"
                $Comment = ""
                if ($CompliancePolicy.securityBlockJailbrokenDevices) {
                    $Status = "Warning"
                    $Comment = "This setting can cause sign in issues."
                    $PolicyWarnings++
                }
                else {
                    $Status = "Supported"
                }
                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    ID                 = $ID
                    PolicyID           = $CompliancePolicy.id
                    PolicyName         = $CompliancePolicy.displayName
                    Setting            = $Setting
                    Value              = $CompliancePolicy.securityBlockJailbrokenDevices
                    TeamsDevicesStatus = $Status
                    Comment            = $Comment 
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceCompliancePolicy')
                $output.Add($SettingPSObj) | Out-Null  


                $ID = 4
                $Setting = "deviceThreatProtectionRequiredSecurityLevel"
                $Comment = ""
                if ($CompliancePolicy.deviceThreatProtectionRequiredSecurityLevel -ne "unavailable") {
                    $Status = "Unsupported"
                    $PolicyErrors++
                }
                else {
                    $Status = "Supported"
                }

                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    ID                 = $ID
                    PolicyID           = $CompliancePolicy.id
                    PolicyName         = $CompliancePolicy.displayName
                    Setting            = $Setting
                    Value              = $CompliancePolicy.deviceThreatProtectionRequiredSecurityLevel
                    TeamsDevicesStatus = $Status 
                    Comment            = $Comment
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceCompliancePolicy')
                $output.Add($SettingPSObj) | Out-Null


                $ID = 5
                $Setting = "securityRequireGooglePlayServices"
                $Comment = ""
                if ($CompliancePolicy.securityRequireGooglePlayServices) {
                    $Status = "Unsupported"
                    $PolicyErrors++
                }
                else {
                    $Status = "Supported"
                }

                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    ID                 = $ID
                    PolicyID           = $CompliancePolicy.id
                    PolicyName         = $CompliancePolicy.displayName
                    Setting            = $SettingPSObj
                    Value              = $CompliancePolicy.securityRequireGooglePlayServices
                    TeamsDevicesStatus = $Status 
                    Comment            = $Comment
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceCompliancePolicy')                            
                $output.Add($SettingPSO) | Out-Null
        

                $ID = 6
                $Setting = "securityRequireUpToDateSecurityProviders"
                $Comment = ""
                if ($CompliancePolicy.securityRequireUpToDateSecurityProviders) {
                    $Status = "Unsupported"
                    $Comment = "Google play isn't installed on Teams Android devices."
                    $PolicyErrors++
                }
                else {
                    $Status = "Supported"
                }

                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    ID                 = $ID
                    PolicyID           = $CompliancePolicy.id
                    PolicyName         = $CompliancePolicy.displayName
                    Setting            = $Setting
                    Value              = $CompliancePolicy.securityRequireUpToDateSecurityProviders
                    TeamsDevicesStatus = $Status 
                    Comment            = $Comment
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceCompliancePolicy')
                $output.Add($SettingPSObj) | Out-Null

        
                $ID = 7
                $Setting = "securityRequireVerifyApps"
                $Comment = ""
                if ($CompliancePolicy.securityRequireVerifyApps) {
                    $Status = "Unsupported"
                    $PolicyErrors++
                }
                else {
                    $Status = "Supported"
                }

                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    ID                 = $ID
                    PolicyID           = $CompliancePolicy.id
                    PolicyName         = $CompliancePolicy.displayName
                    Setting            = $Setting
                    Value              = $CompliancePolicy.securityRequireVerifyApps
                    TeamsDevicesStatus = $Status 
                    Comment            = $Comment
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceCompliancePolicy')
                $output.Add($SettingPSObj) | Out-Null


                $ID = 8.1
                $Setting = "securityRequireSafetyNetAttestationBasicIntegrity"
                $Comment = ""
                if ($CompliancePolicy.securityRequireSafetyNetAttestationBasicIntegrity) {
                    $Status = "Unsupported"
                    $PolicyErrors++
                }
                else {
                    $Status = "Supported"
                }

                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    ID                 = $ID
                    PolicyID           = $CompliancePolicy.id
                    PolicyName         = $CompliancePolicy.displayName
                    Setting            = $Setting
                    Value              = $CompliancePolicy.securityRequireSafetyNetAttestationBasicIntegrity
                    TeamsDevicesStatus = $Status 
                    Comment            = $Comment
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceCompliancePolicy')                            
                $output.Add($SettingPSObj) | Out-Null


                $ID = 8.2
                $Setting = "securityRequireSafetyNetAttestationCertifiedDevice"
                $Comment = ""    
                if ($CompliancePolicy.securityRequireSafetyNetAttestationCertifiedDevice) {
                    $Status = "Unsupported"
                    $PolicyErrors++
                }
                else {
                    $Status = "Supported"
                }

                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    ID                 = $ID
                    PolicyID           = $CompliancePolicy.id
                    PolicyName         = $CompliancePolicy.displayName
                    Setting            = $Setting
                    Value              = $CompliancePolicy.securityRequireSafetyNetAttestationCertifiedDevice
                    TeamsDevicesStatus = $Status 
                    Comment            = $Comment
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceCompliancePolicy')
                $output.Add($SettingPSObj) | Out-Null


                $ID = 9.1
                $Setting = "osMinimumVersion"
                $Comment = ""
                if ($CompliancePolicy.osMinimumVersion) {
                    $Status = "Warning"
                    $Comment = "This setting can cause sign in issues."
                    $PolicyWarnings++
                }
                else {
                    $Status = "Supported"
                }
                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    ID                 = $ID
                    PolicyID           = $CompliancePolicy.id
                    PolicyName         = $CompliancePolicy.displayName
                    Setting            = $Setting
                    Value              = $CompliancePolicy.osMinimumVersion
                    TeamsDevicesStatus = $Status 
                    Comment            = $Comment
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceCompliancePolicy')
                $output.Add($SettingPSObj) | Out-Null  

        
                $ID = 9.2
                $Setting = "osMaximumVersion"
                $Comment = ""
                if ($CompliancePolicy.osMaximumVersion) {
                    $Status = "Warning"
                    $Comment = "This setting can cause sign in issues."
                    $PolicyWarnings++
                }
                else {
                    $Status = "Supported"
                }
                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    ID                 = $ID
                    PolicyID           = $CompliancePolicy.id
                    PolicyName         = $CompliancePolicy.displayName
                    Setting            = $Setting
                    Value              = $CompliancePolicy.osMaximumVersion
                    TeamsDevicesStatus = $Status 
                    Comment            = $Comment
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceCompliancePolicy')
                $output.Add($SettingPSObj) | Out-Null  


                $ID = 10
                $Setting = "storageRequireEncryption"
                $Comment = ""
                if ($CompliancePolicy.storageRequireEncryption) {
                    $Status = "Warning"
                    $Comment = "https://docs.microsoft.com/en-us/microsoftteams/rooms/supported-ca-and-compliance-policies?tabs=phones#supported-device-compliance-policies"
                    $PolicyWarnings++
                }
                else {
                    $Status = "Supported"
                }

                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    ID                 = $ID
                    PolicyID           = $CompliancePolicy.id
                    PolicyName         = $CompliancePolicy.displayName
                    Setting            = $Setting
                    Value              = $CompliancePolicy.storageRequireEncryption
                    TeamsDevicesStatus = $Status 
                    Comment            = $Comment
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceCompliancePolicy')
                $output.Add($SettingPSObj) | Out-Null


                $ID = 11
                $Setting = "securityPreventInstallAppsFromUnknownSources"
                $Comment = ""
                if ($CompliancePolicy.securityPreventInstallAppsFromUnknownSources) {
                    $Status = "Unsupported"
                    $PolicyErrors++
                }
                else {
                    $Status = "Supported"
                }

                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    ID                 = $ID
                    PolicyID           = $CompliancePolicy.id
                    PolicyName         = $CompliancePolicy.displayName
                    Setting            = $Setting
                    Value              = $CompliancePolicy.securityPreventInstallAppsFromUnknownSources
                    TeamsDevicesStatus = $Status 
                    Comment            = $Comment
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceCompliancePolicy')
                $output.Add($SettingPSObj) | Out-Null    



                $ID = 15
                $Setting = "passwordMinutesOfInactivityBeforeLock"
                $Comment = ""
                if ($null -ne $CompliancePolicy.passwordMinutesOfInactivityBeforeLock) {
                    $Status = "Unsupported"
                    $PolicyErrors++
                }
                else {
                    $Status = "Supported"
                }
                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    ID                 = $ID
                    PolicyID           = $CompliancePolicy.id
                    PolicyName         = $CompliancePolicy.displayName
                    Setting            = $Setting
                    Value              = $CompliancePolicy.passwordMinutesOfInactivityBeforeLock
                    TeamsDevicesStatus = $Status 
                    Comment            = $Comment
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceCompliancePolicy')
                $output.Add($SettingPSObj) | Out-Null  


                $ID = 16
                $Setting = "passwordRequired"
                $Comment = ""
                if ($CompliancePolicy.passwordRequired) {
                    $Status = "Unsupported"
                    $PolicyErrors++
                }
                else {
                    $Status = "Supported"
                }
                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    ID                 = $ID
                    PolicyID           = $CompliancePolicy.id
                    PolicyName         = $CompliancePolicy.displayName
                    Setting            = $Setting
                    Value              = $CompliancePolicy.passwordRequired
                    TeamsDevicesStatus = $Status 
                    Comment            = $Comment
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceCompliancePolicy')
                $output.Add($SettingPSObj) | Out-Null  


                $ID = 17.1
                $Setting = "passwordRequiredType"
                $Comment = ""
                if ($CompliancePolicy.passwordRequiredType -ne 'deviceDefault') {
                    $Status = "Unsupported"
                    $PolicyErrors++
                }
                else {
                    $Status = "Supported"
                }
                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    ID                 = $ID
                    PolicyID           = $CompliancePolicy.id
                    PolicyName         = $CompliancePolicy.displayName
                    Setting            = $Setting
                    Value              = $CompliancePolicy.passwordRequiredType
                    TeamsDevicesStatus = $Status 
                    Comment            = $Comment
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceCompliancePolicy')
                $output.Add($SettingPSObj) | Out-Null  


                $ID = 17.2
                $Setting = "passwordMinimumLength"
                $Comment = ""
                if ($null -ne $CompliancePolicy.passwordMinimumLength) {
                    $Status = "Unsupported"
                    $PolicyErrors++
                }
                else {
                    $Status = "Supported"
                }
                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    ID                 = $ID
                    PolicyID           = $CompliancePolicy.id
                    PolicyName         = $CompliancePolicy.displayName
                    Setting            = $Setting
                    Value              = $CompliancePolicy.passwordMinimumLength
                    TeamsDevicesStatus = $Status 
                    Comment            = $Comment
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceCompliancePolicy')
                $output.Add($SettingPSObj) | Out-Null  

        
                $ID = 17.3
                $Setting = "passwordExpirationDays"
                $Comment = ""
                if ($null -ne $CompliancePolicy.passwordExpirationDays) {
                    $Status = "Unsupported"
                    $PolicyErrors++
                }
                else {
                    $Status = "Supported"
                }
                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    ID                 = $ID
                    PolicyID           = $CompliancePolicy.id
                    PolicyName         = $CompliancePolicy.displayName
                    Setting            = $Setting
                    Value              = $CompliancePolicy.passwordExpirationDays 
                    TeamsDevicesStatus = $Status 
                    Comment            = $Comment
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceCompliancePolicy')
                $output.Add($SettingPSObj) | Out-Null  


                $ID = 17.4
                $Setting = "passwordPreviousPasswordBlockCount"
                $Comment = ""
                if ($null -ne $CompliancePolicy.passwordPreviousPasswordBlockCount) {
                    $Status = "Unsupported"
                    $PolicyErrors++
                }
                else {
                    $Status = "Supported"
                }
                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    ID                 = $ID
                    PolicyID           = $CompliancePolicy.id
                    PolicyName         = $CompliancePolicy.displayName
                    Setting            = $Setting
                    Value              = $CompliancePolicy.passwordPreviousPasswordBlockCount
                    TeamsDevicesStatus = $Status 
                    Comment            = $Comment
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceCompliancePolicy')
                $output.Add($SettingPSObj) | Out-Null  


                $ID = 18
                $Setting = "minAndroidSecurityPatchLevel"
                $Comment = ""
                if ($CompliancePolicy.minAndroidSecurityPatchLevel -ne "") {
                    $Status = "Warning"
                    $Comment = "This setting can cause sign in issues."
                    $PolicyWarnings++
                }
                else {
                    $Status = "Supported"
                }

                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    ID                 = $ID
                    PolicyID           = $CompliancePolicy.id
                    PolicyName         = $CompliancePolicy.displayName
                    Setting            = $Setting
                    Value              = $CompliancePolicy.minAndroidSecurityPatchLevel
                    TeamsDevicesStatus = $Status 
                    Comment            = $Comment
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceCompliancePolicy')
                $output.Add($SettingPSObj) | Out-Null

                if ($PolicyErrors -gt 0) {
                    $StatusSum = "Unsupported"
                }
                elseif ($PolicyWarnings -gt 0) {
                    $StatusSum = "Warning"
                }
                else {
                    $StatusSum = "Supported"
                }

                $PolicySum = New-Object -TypeName PSObject -Property @{
                    PolicyID           = $CompliancePolicy.id
                    PolicyName         = $CompliancePolicy.displayName
                    TeamsDevicesStatus = $StatusSum
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceCompliancePolicySumary')
                $outputSum.Add($PolicySum) | Out-Null
            }
        }
        if ($Detailed) {
            $output | Sort-Object PolicyName, ID 
        }
        else {
            $outputSum 
        }
    }
}