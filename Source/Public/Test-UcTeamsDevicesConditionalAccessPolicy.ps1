<#
.SYNOPSIS
Validate which Conditional Access policies are supported by Microsoft Teams Android Devices

.DESCRIPTION
This function will validate each setting in a Conditional Access Policy to make sure they are in line with the supported settings:

    https://docs.microsoft.com/microsoftteams/rooms/supported-ca-and-compliance-policies?tabs=phones#conditional-access-policies"

Contributors: Traci Herr, David Paulino

Requirements: Microsoft Graph PowerShell Module (Install-Module Microsoft.Graph)

.PARAMETER Detailed
Displays test results for all settings in each Conditional Access Policy

.EXAMPLE 
PS> Test-UcTeamsDevicesConditionalAccessPolicy

.EXAMPLE 
PS> Test-UcTeamsDevicesConditionalAccessPolicy -Detailed

#>

Function Test-UcTeamsDevicesConditionalAccessPolicy {

    Param(
        [Parameter(Mandatory = $false)]
        [switch]$Detailed
    )

    $connectedMSGraph = $false
    $ConditionalAccessPolicies = $null
    $URLTeamsDevicesCA = "aka.ms/TeamsDevicesAndroidPolicies#supported-conditional-access-policies"
    $URLTeamsDevicesKnownIssues = "https://docs.microsoft.com/microsoftteams/troubleshoot/teams-rooms-and-devices/rooms-known-issues#teams-phone-devices"

    $scopes = (Get-MgContext).Scopes

    if (!($scopes) -or !( "Policy.Read.All" -in $scopes )) {
        Connect-MgGraph -Scopes "Policy.Read.All"
    }

    try {
        $ConditionalAccessPolicies = (Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -Method GET).Value
        $connectedMSGraph = $true
    }
    catch {
        Write-Error 'Please connect to MS Graph with Connect-MgGraph -Scopes "Policy.Read.All" before running this script'
    }

    if ($connectedMSGraph) {
        $output = [System.Collections.ArrayList]::new()
        $outputSum = [System.Collections.ArrayList]::new()
        foreach ($ConditionalAccessPolicy in $ConditionalAccessPolicies) {

            $CAPolicyState = $ConditionalAccessPolicy.State

            if ($CAPolicyState -eq "enabledForReportingButNotEnforced") {
                $CAPolicyState = "ReportOnly"
            }

            $PolicyErrors = 0
            $PolicyWarnings = 0

            #Cloud Apps
            #Exchange 00000002-0000-0ff1-ce00-000000000000
            #SharePoint 00000003-0000-0ff1-ce00-000000000000
            #Teams cc15fd57-2c6c-4117-a88c-83b1d56b4bbe
            $hasExchange = $false
            $hasSharePoint = $false
            $hasTeams = $false
            $hasOffice365 = $false
            $CloudAppValue = ""
            foreach ($Application in $ConditionalAccessPolicy.Conditions.Applications.IncludeApplications) {

                switch ($Application) {
                    "All" { $hasOffice365 = $true; $CloudAppValue = "All" }
                    "Office365" { $hasOffice365 = $true; $CloudAppValue = "Office 365" }
                    "00000002-0000-0ff1-ce00-000000000000" { $hasExchange = $true; $CloudAppValue += "Exchange, " }
                    "00000003-0000-0ff1-ce00-000000000000" { $hasSharePoint = $true; $CloudAppValue += "SharePoint, " }
                    "cc15fd57-2c6c-4117-a88c-83b1d56b4bbe" { $hasTeams = $true; $CloudAppValue += "Teams, " }
                    "None" { $CloudAppValue = "None" }
                }
            }
            if ($CloudAppValue.EndsWith(", ")) {
                $CloudAppValue = $CloudAppValue.Substring(0, $CloudAppValue.Length - 2)
            }

            if (($hasExchange -and $hasSharePoint -and $hasTeams) -or ($hasOffice365)) {
                $Status = "Supported"
                $Comment = ""
            }
            else {
                $Status = "Unsupported"
                $Comment = "Teams Devices needs to access: Office 365 or Exchange Online, SharePoint Online, and Microsoft Teams"
                $PolicyErrors++
            }
            
            $SettingPSObj = New-Object -TypeName PSObject -Property @{
                CAPolicyID         = $ConditionalAccessPolicy.id
                CAPolicyName       = $ConditionalAccessPolicy.displayName
                CAPolicyState      = $CAPolicyState
                Setting            = "Cloud Apps or actions"
                Value              = $CloudAppValue
                TeamsDevicesStatus = $Status 
                Comment            = $Comment
            }
            $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceConditionalAccessPolicy')
            $output.Add($SettingPSObj) | Out-Null

            #Conditions - ClientAppTypes
            foreach ($ClientAppType in $ConditionalAccessPolicy.Conditions.ClientAppTypes) {
                if ($ClientAppType -eq "All") {
                    $Status = "Supported"
                    $Comment = ""
                }
                else {
                    $PolicyErrors++
                    $Status = "Unsupported"
                    $Comment = $URLTeamsDevicesCA
                }

                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    CAPolicyID         = $ConditionalAccessPolicy.id
                    CAPolicyName       = $ConditionalAccessPolicy.displayName
                    CAPolicyState      = $CAPolicyState
                    Setting            = "Conditions - ClientAppTypes"
                    Value              = $ClientAppType
                    TeamsDevicesStatus = $Status 
                    Comment            = $Comment
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceConditionalAccessPolicy')
                $output.Add($SettingPSObj) | Out-Null
            }

            #Device

            if ($ConditionalAccessPolicy.conditions.devices) {
                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    CAPolicyID         = $ConditionalAccessPolicy.id
                    CAPolicyName       = $ConditionalAccessPolicy.displayName
                    CAPolicyState      = $CAPolicyState
                    Setting            = "Device Filter"
                    Value              = "Configured"
                    TeamsDevicesStatus = "Supported"
                    Comment            = $ConditionalAccessPolicy.conditions.devices.deviceFilter.mode + ": " + $ConditionalAccessPolicy.conditions.devices.deviceFilter.rule
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceConditionalAccessPolicy')
                $output.Add($SettingPSObj) | Out-Null
            }
            else {
                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    CAPolicyID         = $ConditionalAccessPolicy.id
                    CAPolicyName       = $ConditionalAccessPolicy.displayName
                    CAPolicyState      = $CAPolicyState
                    Setting            = "Device Filter"
                    Value              = "Missing"
                    TeamsDevicesStatus = "Warning"
                    Comment            = "Device Filter is required when multiple Conditional Access policies exist."
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceConditionalAccessPolicy')
                $output.Add($SettingPSObj) | Out-Null
            }



            #Grant Controls

            $Setting = "GrantControls"
            foreach ($BuiltInControl in $ConditionalAccessPolicy.GrantControls.BuiltInControls) {
                $Comment = "" 
                if ($BuiltInControl -in 'DomainJoinedDevice', 'ApprovedApplication', 'CompliantApplication', 'PasswordChange') {
                    $PolicyErrors++
                    $Status = "Unsupported"
                    $Comment = $URLTeamsDevicesCA
                }
                else {
                    $Status = "Supported"
                }

                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    CAPolicyID         = $ConditionalAccessPolicy.id
                    CAPolicyName       = $ConditionalAccessPolicy.displayName
                    CAPolicyState      = $CAPolicyState
                    Setting            = $Setting 
                    Value              = $BuiltInControl 
                    TeamsDevicesStatus = $Status 
                    Comment            = $Comment
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceConditionalAccessPolicy')
                $output.Add($SettingPSObj) | Out-Null
            }

            
            if ($ConditionalAccessPolicy.GrantControls.CustomAuthenticationFactors) {
                $PolicyWarnings++
                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    CAPolicyID         = $ConditionalAccessPolicy.id
                    CAPolicyName       = $ConditionalAccessPolicy.displayName
                    CAPolicyState      = $CAPolicyState
                    Setting            = $Setting 
                    Value              = "CustomAuthenticationFactors"
                    TeamsDevicesStatus = "Unsupported"
                    Comment            = $URLTeamsDevicesCA
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceConditionalAccessPolicy')
                $output.Add($SettingPSObj) | Out-Null
            }

            if ($ConditionalAccessPolicy.GrantControls.TermsOfUse) {
                $PolicyWarnings++
                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    CAPolicyID         = $ConditionalAccessPolicy.id
                    CAPolicyName       = $ConditionalAccessPolicy.displayName
                    CAPolicyState      = $CAPolicyState
                    Setting            = $Setting 
                    Value              = "Terms of Use"
                    TeamsDevicesStatus = "Warning"
                    Comment            = $URLTeamsDevicesKnownIssues
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceConditionalAccessPolicy')
                $output.Add($SettingPSObj) | Out-Null
            }

            $Setting = "SessionControls"
            $Comment = "" 
            if ($ConditionalAccessPolicy.SessionControls.ApplicationEnforcedRestrictions) {
                $PolicyErrors++
                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    CAPolicyID         = $ConditionalAccessPolicy.id
                    CAPolicyName       = $ConditionalAccessPolicy.displayName
                    CAPolicyState      = $CAPolicyState
                    Setting            = $Setting 
                    Value              = "ApplicationEnforcedRestrictions"
                    TeamsDevicesStatus = "Unsupported" 
                    Comment            = $Comment
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceConditionalAccessPolicy')
                $output.Add($SettingPSObj) | Out-Null
            }
            if ($ConditionalAccessPolicy.SessionControls.CloudAppSecurity) {
                $PolicyErrors++
                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    CAPolicyID         = $ConditionalAccessPolicy.id
                    CAPolicyName       = $ConditionalAccessPolicy.displayName
                    CAPolicyState      = $CAPolicyState
                    Setting            = $Setting 
                    Value              = "CloudAppSecurity"
                    TeamsDevicesStatus = "Unsupported" 
                    Comment            = $URLTeamsDevicesCA
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceConditionalAccessPolicy')
                $output.Add($SettingPSObj) | Out-Null
            }
            if ($ConditionalAccessPolicy.SessionControls.SignInFrequency) {
                $PolicyWarnings++
                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    CAPolicyID         = $ConditionalAccessPolicy.id
                    CAPolicyName       = $ConditionalAccessPolicy.displayName
                    CAPolicyState      = $CAPolicyState
                    Setting            = $Setting 
                    Value              = "SignInFrequency"
                    TeamsDevicesStatus = "Warning" 
                    Comment            = "Users will be signout from Teams Device every " + $ConditionalAccessPolicy.SessionControls.SignInFrequency.Value + " " + $ConditionalAccessPolicy.SessionControls.SignInFrequency.Type
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceConditionalAccessPolicy')
                $output.Add($SettingPSObj) | Out-Null
            }
            if ($ConditionalAccessPolicy.SessionControls.PersistentBrowser) {
                $PolicyErrors++
                $SettingPSObj = New-Object -TypeName PSObject -Property @{
                    CAPolicyID         = $ConditionalAccessPolicy.id

                    CAPolicyName       = $ConditionalAccessPolicy.displayName
                    CAPolicyState      = $CAPolicyState
                    Setting            = $Setting 
                    Value              = "PersistentBrowser"
                    TeamsDevicesStatus = "Unsupported" 
                    Comment            = $URLTeamsDevicesCA
                }
                $SettingPSObj.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceConditionalAccessPolicy')
                $output.Add($SettingPSObj) | Out-Null
            }

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
                CAPolicyID         = $ConditionalAccessPolicy.id
                CAPolicyName       = $ConditionalAccessPolicy.DisplayName
                CAPolicyState      = $CAPolicyState
                TeamsDevicesStatus = $StatusSum
            }
            $PolicySum.PSObject.TypeNAmes.Insert(0, 'TeamsDeviceConditionalAccessPolicySummary')
            $outputSum.Add($PolicySum) | Out-Null
        }
        if ($Detailed) {
            $output | Sort-Object CAPolicyName, Setting
        }
        else {
            $outputSum 
        }
    }
}