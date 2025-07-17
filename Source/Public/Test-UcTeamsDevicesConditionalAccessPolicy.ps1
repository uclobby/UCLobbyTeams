function Test-UcTeamsDevicesConditionalAccessPolicy {
    <#
        .SYNOPSIS
        Validate which Conditional Access policies are supported by Microsoft Teams Android Devices

        .DESCRIPTION
        This function will validate each setting in a Conditional Access Policy to make sure they are in line with the supported settings:

            https://docs.microsoft.com/microsoftteams/rooms/supported-ca-and-compliance-policies?tabs=phones#conditional-access-policies"

        Contributors: Traci Herr, David Paulino, Gonçalo Sepúlveda and Miguel Ferreira

        Requirements:   EntraAuth PowerShell Module (Install-Module EntraAuth)
                        or
                        Microsoft Graph Authentication PowerShell Module (Install-Module Microsoft.Graph.Authentication)

                        Microsoft Graph Scopes:
                            "Policy.Read.All", "Directory.Read.All"

        .PARAMETER PolicyID
        Specifies a Policy ID that will be checked if is supported by Microsoft Teams Devices.
        
        .PARAMETER UserUPN
        Specifies a UserUPN that we want to check for applied Conditional Access policies.

        .PARAMETER DeviceType
        Type of Teams Device:
            MTRWindows - Microsoft Teams Room Running Windows
            MTRAndroidAndPanel - Microsoft Teams Room Running Android and Panels
            PhoneAndDisplay - Microsoft Teams Phone and Displays 

        .PARAMETER Detailed
        Displays test results for all settings in each Conditional Access Policy.

        .PARAMETER All
        When present it check all Conditional Access policies independently if they are assigned to a User(s)/Group(s).

        .PARAMETER IncludeSupported
        Displays results for all settings in each Conditional Access Policy.

        .PARAMETER ExportCSV
        When present will export the detailed results to a CSV file. By default, will save the file under the current user downloads, unless we specify the OutputPath.

        .PARAMETER OutputPath
        Allows to specify the path where we want to save the results.

        .EXAMPLE 
        PS> Test-UcTeamsDevicesConditionalAccessPolicy

        .EXAMPLE 
        PS> Test-UcTeamsDevicesConditionalAccessPolicy -All

        .EXAMPLE 
        PS> Test-UcTeamsDevicesConditionalAccessPolicy -Detailed

        .EXAMPLE 
        PS> Test-UcTeamsDevicesConditionalAccessPolicy -Detailed -IncludedSupported

        .EXAMPLE 
        PS> Test-UcTeamsDevicesConditionalAccessPolicy -UserUPN
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [string]$PolicyID,
        [string]$UserUPN,
        [ValidateSet("MTRWindows", "MTRAndroidAndPanel", "PhoneAndDisplay")]
        [string]$DeviceType,
        [switch]$Detailed,
        [switch]$All,
        [switch]$IncludeSupported,
        [switch]$ExportCSV,
        [string]$OutputPath
    )
    $GraphPathUsers = "/users"
    $GraphPathGroups = "/groups"
    $GraphPathConditionalAccess = "/identity/conditionalAccess/policies"

    $URLTeamsDevicesCA = "https://aka.ms/TeamsDevicePolicies#supported-conditional-access-policies"
    $URLTeamsDevicesKnownIssues = "https://docs.microsoft.com/microsoftteams/troubleshoot/teams-rooms-and-devices/rooms-known-issues#issues-with-teams-phones"

    $output = [System.Collections.ArrayList]::new()
    $outputSum = [System.Collections.ArrayList]::new()
    $Groups = New-Object 'System.Collections.Generic.Dictionary[string, string]'

    #region Graph Connection, Scope validation and module version
    if (!(Test-UcServiceConnection -Type MSGraph -Scopes "Policy.Read.All", "Directory.Read.All")) {
        return
    }
    if (!($global:UCLobbyTeamsModuleCheck)) {
        Test-UcPowerShellModule -ModuleName UcLobbyTeams | Out-Null
        $global:UCLobbyTeamsModuleCheck = $true
    }
    #endregion

    #region ExportCSV defaults to Detaileld and Output Path Test
    if ($ExportCSV) {
        $Detailed = $true
        $outFileName = "TeamsDevices_ConditionalAccessPolicy_Report_" + ( get-date ).ToString('yyyyMMdd-HHmmss') + ".csv"
        if ($OutputPath) {
            if (!(Test-Path $OutputPath -PathType Container)) {
                Write-Host ("Error: Invalid folder " + $OutputPath) -ForegroundColor Red
                return
            } 
            $OutputFullPath = [System.IO.Path]::Combine($OutputPath, $outFileName)
        }
        else {                
            $OutputFullPath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads", $outFileName)
        }
    }
    #endregion

    try {
        Write-Progress -Activity "Test-UcTeamsDevicesConditionalAccessPolicy" -Status "Getting Conditional Access Policies"
        if ($PolicyID) {
            $ConditionalAccessPolicies = Invoke-UcGraphRequest -Path ($GraphPathConditionalAccess + "/" + $PolicyID) -Beta
        }
        else {
            $ConditionalAccessPolicies = Invoke-UcGraphRequest -Path ($GraphPathConditionalAccess) -Beta
        }
    }
    catch [System.Net.Http.HttpRequestException] {
        if ($PSItem.Exception.Response.StatusCode -eq "Forbidden") {
            Write-Host "Access Denied, please make sure the user connecing to MS Graph is part of one of the following Global Reader/Conditional Access Administrator/Global Administrator roles"
            return
        }
        else {
            Write-Error $PSItem.Exception.Message
        }
    }
    catch {
        Write-Error $PSItem.Exception.Message
    }

    if ($UserUPN) {
        try {
            $UserID = (Invoke-UcGraphRequest -Path ($GraphPathUsers + "/" + $UserUPN + "?`$select=id")).id
            $UserGroups = (Invoke-UcGraphRequest -Path ($GraphPathUsers + "/" + $UserUPN + "/transitiveMemberOf?`$select=id")).id
        }
        catch [System.Net.Http.HttpRequestException] {
            if ($PSItem.Exception.Response.StatusCode -eq "NotFound") {
                Write-Warning -Message ("User Not Found: " + $UserUPN)
            }
            return
        }
    }

    try {
        Write-Progress -Activity "Test-UcTeamsDevicesConditionalAccessPolicy" -Status "Fetching Service Principals details."
        $ServicePrincipals = Invoke-UcGraphRequest -Path "/servicePrincipals?`$select=AppId,DisplayName" 
    }
    catch {}

    $p = 0
    $policyCount = $ConditionalAccessPolicies.Count
    foreach ($ConditionalAccessPolicy in $ConditionalAccessPolicies) {
        $p++
        Write-Progress -Activity "Test-UcTeamsDevicesConditionalAccessPolicy" -Status ("Checking policy " + $ConditionalAccessPolicy.displayName + " - $p of $policyCount")
        $AssignedToGroup = [System.Collections.ArrayList]::new()
        $ExcludedFromGroup = [System.Collections.ArrayList]::new()
        $AssignedToUserCount = 0
        $ExcludedFromUserCount = 0
        $outAssignedToGroup = ""
        $outExcludedFromGroup = ""
        $userIncluded = $false
        $userExcluded = $false
        $AccessType = "Unknown"
        $StatusSum = ""
    
        $totalCAPolicies++
        $PolicyErrors = 0
        $PolicyWarnings = 0
        
        if ($UserUPN) {
            if ($UserID -in $ConditionalAccessPolicy.conditions.users.excludeUsers) {
                $userExcluded = $true
                Write-Verbose -Message ("Skiping conditional access policy " + $ConditionalAccessPolicy.displayName + ", since user " + $UserUPN + " is part of Excluded Users")
            }
            elseif ($UserID -in $ConditionalAccessPolicy.conditions.users.includeUsers) {
                $userIncluded = $true
            }
        }  

        #All Users in Conditional Access Policy will show as a 'All' in the includeUsers.
        if ("All" -in $ConditionalAccessPolicy.conditions.users.includeUsers) {
            $GroupEntry = New-Object -TypeName PSObject -Property @{
                GroupID          = "All"
                GroupDisplayName = "All Users"
            }
            [void]$AssignedToGroup.Add($GroupEntry)
            $userIncluded = $true
        }
        elseif ((($ConditionalAccessPolicy.conditions.users.includeUsers).count -gt 0) -and "None" -notin $ConditionalAccessPolicy.conditions.users.includeUsers) {
            $AssignedToUserCount = ($ConditionalAccessPolicy.conditions.users.includeUsers).count
            if (!$UserUPN) {
                $userIncluded = $true
            }
            foreach ($includedGroup in $ConditionalAccessPolicy.conditions.users.includeGroups) {
                $GroupDisplayName = $includedGroup
                if ($Groups.ContainsKey($includedGroup)) {
                    $GroupDisplayName = $Groups.Item($includedGroup)
                }
                else {
                    try {
                        $GroupInfo = Invoke-UcGraphRequest -Path ($GraphPathGroups + "/" + $includedGroup + "/?`$select=id,displayname")
                        $Groups.Add($GroupInfo.id, $GroupInfo.displayname)
                        $GroupDisplayName = $GroupInfo.displayname
                    }
                    catch {
                    }
                }
                $GroupEntry = New-Object -TypeName PSObject -Property @{
                    GroupID          = $includedGroup
                    GroupDisplayName = $GroupDisplayName
                }
                #2025/03/21:  We add the group and only flag userincluded if the group is part of the groups that the user is part of.
                $AssignedToGroup.Add($GroupEntry) | Out-Null
                if ($includedGroup -in $UserGroups) {
                    $userIncluded = $true
                }
            }
        }

        foreach ($excludedGroup in $ConditionalAccessPolicy.conditions.users.excludeGroups) {
            $GroupDisplayName = $excludedGroup
            if ($Groups.ContainsKey($excludedGroup)) {
                $GroupDisplayName = $Groups.Item($excludedGroup)
            }
            else {
                try {
                    $GroupInfo = Invoke-UcGraphRequest -Path ($GraphPathGroups + "/" + $excludedGroup + "/?`$select=id,displayname")
                    $Groups.Add($GroupInfo.id, $GroupInfo.displayname)
                    $GroupDisplayName = $GroupInfo.displayname
                }
                catch { }
            }
            $GroupEntry = New-Object -TypeName PSObject -Property @{
                GroupID          = $excludedGroup
                GroupDisplayName = $GroupDisplayName
            }
            $ExcludedFromGroup.Add($GroupEntry) | Out-Null                
            if ($excludedGroup -in $UserGroups) {
                $userExcluded = $true
                Write-Verbose ("Skiping conditional access policy " + $ConditionalAccessPolicy.displayName + ", since user " + $UserUPN + " is part of an Excluded Group: " + $GroupEntry.GroupDisplayName)
            }
        }
        $ExcludedFromUserCount = ($ConditionalAccessPolicy.conditions.users.excludeUsers).count

        if ("GuestsOrExternalUsers" -in $ConditionalAccessPolicy.conditions.users.excludeUsers) {
            $ExcludedFromUserCount--
        }

        #If only assigned/excluded from a group we will show the group display name, otherwise the number of groups assigned/excluded.
        if (($AssignedToGroup.count -gt 0) -and ($AssignedToUserCount -gt 0)) {
            $outAssignedToGroup = "$AssignedToUserCount user(s)," + $AssignedToGroup.count + " group(s)"
        }
        elseif (($AssignedToGroup.count -eq 0) -and ($AssignedToUserCount -gt 0)) {
            $outAssignedToGroup = "$AssignedToUserCount user(s)"
        }
        elseif (($AssignedToGroup.count -gt 0) -and ($AssignedToUserCount -eq 0)) {
            if ($AssignedToGroup.count -eq 1) {
                $outAssignedToGroup = $AssignedToGroup[0].GroupDisplayName
            }
            else {
                $outAssignedToGroup = "" + $AssignedToGroup.count + " group(s)"
            }
        }
        else {
            $outAssignedToGroup = "None"
        }

        if (($ExcludedFromGroup.count -gt 0) -and ($ExcludedFromUserCount -gt 0)) {
            $outExcludedFromGroup = "$ExcludedFromUserCount user(s), " + $ExcludedFromGroup.count + " group(s)"
        }
        elseif (($ExcludedFromGroup.count -eq 0) -and ($ExcludedFromUserCount -gt 0)) {
            $outExcludedFromGroup = "$ExcludedFromUserCount user(s)"
        }
        elseif (($ExcludedFromGroup.count -gt 0) -and ($ExcludedFromUserCount -eq 0)) {
            if ($ExcludedFromGroup.count -eq 1) {
                $outExcludedFromGroup = $ExcludedFromGroup[0].GroupDisplayName
            }
            else {
                $outExcludedFromGroup = "" + $ExcludedFromGroup.count + " group(s)"
            }
        }
        else {
            $outExcludedFromGroup = "None"
        }

        $PolicyState = $ConditionalAccessPolicy.State
        if ($PolicyState -eq "enabledForReportingButNotEnforced") {
            $PolicyState = "ReportOnly"
        }
                
        #region 2: Assignment > Cloud apps or actions > Cloud Apps
        #Exchange 00000002-0000-0ff1-ce00-000000000000
        #SharePoint 00000003-0000-0ff1-ce00-000000000000
        #Teams cc15fd57-2c6c-4117-a88c-83b1d56b4bbe
        $ID = 2
        $Setting = "CloudApps"
        $SettingDescription = "Assignment > Cloud apps or actions > Cloud Apps"
        $Comment = ""
        $hasExchange = $false
        $hasSharePoint = $false
        $hasTeams = $false
        $hasOffice365 = $false
        #2025-06-18: Checking if policy is targeted to Intune Enrollment Service
        $hasIntune = $false
        $SettingValue = ""
        foreach ($Application in $ConditionalAccessPolicy.Conditions.Applications.IncludeApplications) {
            $appDisplayName = ($ServicePrincipals |  Where-Object -Property AppId -eq -Value $Application).DisplayName
            switch ($Application) {
                "All" { $hasOffice365 = $true; $hasIntune = $true; $SettingValue = "All" }
                "Office365" { $hasOffice365 = $true; $SettingValue = "Office 365" }
                "00000002-0000-0ff1-ce00-000000000000" { $hasExchange = $true; $SettingValue += $appDisplayName + "; " }
                "00000003-0000-0ff1-ce00-000000000000" { $hasSharePoint = $true; $SettingValue += $appDisplayName + "; " }
                "cc15fd57-2c6c-4117-a88c-83b1d56b4bbe" { $hasTeams = $true; $SettingValue += $appDisplayName + "; " }
                "d4ebce55-015a-49b5-a083-c84d1797ae8c" { $hasIntune = $true; }
                default { $SettingValue += $appDisplayName + "; " }
            }
        }
        if ($SettingValue.EndsWith("; ")) {
            $SettingValue = $SettingValue.Substring(0, $SettingValue.Length - 2)
        }

        #2025-03-21: We need to take into consideration if it's a block or allow policy.
        $isBlockAccessCA = $ConditionalAccessPolicy.GrantControls.BuiltInControls -contains "block"
        if ($isBlockAccessCA) {
            $AccessType = "Block"
        }
        elseif ($ConditionalAccessPolicy.GrantControls.BuiltInControls -or $ConditionalAccessPolicy.SessionControls -or $ConditionalAccessPolicy.GrantControls.authenticationStrength) {
            $AccessType = "Allow"
        } 

        #region 2025-03-25: Adding support for Device Platform 
        $includeDevicePlatform = $true
        if ($ConditionalAccessPolicy.conditions.platforms) {
            $includeDevicePlatform = $false
            switch ($DeviceType) {
                "MTRWindows" { $TeamsDevicesPlatForms = "windows" }
                "PhoneAndDisplay" { $TeamsDevicesPlatForms = "android" }
                "MTRAndroidAndPanel" { $TeamsDevicesPlatForms = "android" } 
                default { $TeamsDevicesPlatForms = "windows", "android" }
            }
            foreach ($TeamsDevicesPlatForm in $TeamsDevicesPlatForms) {
                if (($ConditionalAccessPolicy.conditions.platforms.includePlatforms -eq "all" -or $TeamsDevicesPlatForm -in $ConditionalAccessPolicy.conditions.platforms.includePlatforms) -and !($TeamsDevicesPlatForm -in $ConditionalAccessPolicy.conditions.platforms.excludePlatforms)) {
                    $includeDevicePlatform = $true
                }
            }
        }
        #endregion
        #region 2025-03-25: Skipping ClientApps and Block
        $includeClientApps = $false
        if ($ConditionalAccessPolicy.Conditions.ClientAppTypes | Where-Object { $_ -in ("all", "mobileAppsAndDesktopClients", "browser") } ) {
            $includeClientApps = $true
        }
        #endregion
        Write-Verbose ("Checking if Policy " + $ConditionalAccessPolicy.DisplayName + " is assigned: `
                            Policy State $PolicyState" + ", `
                            Assigned Group " + $AssignedToGroup.count + ", `
                            has Office365/Teams " + $hasOffice365 + "/" + $hasTeams + ", `
                            User Included/Excluded " + $userIncluded + "/" + $userExcluded + ", `
                            Device Platform " + ($TeamsDevicesPlatForms -join ",")) 
        if ((((($AssignedToGroup.count -gt 0) `
                            -or $userIncluded) `
                        -and ($hasOffice365 `
                            -or $hasTeams `
                            -or $hasExchange `
                            -or $hasSharePoint `
                            -or $hasIntune) `
                        -and ($PolicyState -NE "disabled") `
                        -and $includeDevicePlatform `
                        -and $includeClientApps) `
                    -and (!$userExcluded)) `
                -or $all) {
            #2025-03-25: Supported if the required target resources are present and not a blocking Conditional Access
            if ((($hasExchange -and $hasSharePoint -and $hasTeams) -or ($hasOffice365)) -and !$isBlockAccessCA) {
                $Status = "Supported"
            }
            else {
                $Status = "Unsupported"
                $Comment = "Teams Devices needs to access: Office 365 or Exchange Online, SharePoint Online, and Microsoft Teams"
                $PolicyErrors++
            }
            $SettingPSObj = [PSCustomObject][Ordered]@{
                ID                    = $ID
                PolicyName            = $ConditionalAccessPolicy.displayName
                PolicyID              = $ConditionalAccessPolicy.id
                PolicyState           = $PolicyState
                PolicyAccessType      = $AccessType
                AssignedToGroup       = $outAssignedToGroup
                AssignedToGroupList   = $AssignedToGroup
                ExcludedFromGroup     = $outExcludedFromGroup 
                ExcludedFromGroupList = $ExcludedFromGroup
                TeamsDevicesStatus    = $Status 
                Setting               = $Setting 
                SettingDescription    = $SettingDescription 
                Value                 = $SettingValue
                Comment               = $Comment
            }
            [void]$output.Add($SettingPSObj)
            #endregion
            
            #region 6: Conditions > Locations
            $ID = 6.1
            $Setting = "includeLocations"
            $SettingDescription = "Conditions > Locations"
            $SettingValue = "Not Configured"
            $Comment = "" 
            $Status = "Supported"
            if ($ConditionalAccessPolicy.conditions.locations.includeLocations) {
                $SettingValue = $ConditionalAccessPolicy.conditions.locations.includeLocations
            }
            $SettingPSObj = [PSCustomObject][Ordered]@{
                ID                    = $ID
                PolicyName            = $ConditionalAccessPolicy.displayName
                PolicyID              = $ConditionalAccessPolicy.id
                PolicyState           = $PolicyState
                PolicyAccessType      = $AccessType
                AssignedToGroup       = $outAssignedToGroup
                AssignedToGroupList   = $AssignedToGroup
                ExcludedFromGroup     = $outExcludedFromGroup 
                ExcludedFromGroupList = $ExcludedFromGroup
                TeamsDevicesStatus    = $Status 
                Setting               = $Setting 
                SettingDescription    = $SettingDescription 
                Value                 = $SettingValue
                Comment               = $Comment
            }
            [void]$output.Add($SettingPSObj)
            
            $ID = 6.2
            $Setting = "excludeLocations"
            $SettingDescription = "Conditions > Locations"
            $SettingValue = "Not Configured"
            $Comment = "" 
            $Status = "Supported"
            if ($ConditionalAccessPolicy.conditions.locations.excludeLocations) {
                $SettingValue = $ConditionalAccessPolicy.conditions.locations.excludeLocations
            }
            $SettingPSObj = [PSCustomObject][Ordered]@{
                ID                    = $ID
                PolicyName            = $ConditionalAccessPolicy.displayName
                PolicyID              = $ConditionalAccessPolicy.id
                PolicyState           = $PolicyState
                PolicyAccessType      = $AccessType
                AssignedToGroup       = $outAssignedToGroup
                AssignedToGroupList   = $AssignedToGroup
                ExcludedFromGroup     = $outExcludedFromGroup 
                ExcludedFromGroupList = $ExcludedFromGroup
                TeamsDevicesStatus    = $Status 
                Setting               = $Setting 
                SettingDescription    = $SettingDescription 
                Value                 = $SettingValue
                Comment               = $Comment
            }
            [void]$output.Add($SettingPSObj)
            #endregion
            
            #region 7: Conditions > Client apps
            $ID = 7
            $Setting = "ClientAppTypes"
            $SettingDescription = "Conditions > Client apps"
            $SettingValue = ""
            $Comment = ""
            foreach ($ClientAppType in $ConditionalAccessPolicy.Conditions.ClientAppTypes) {
                if ($ClientAppType -eq "All") {
                    $Status = "Supported"
                    $SettingValue = $ClientAppType
                }
                else {
                    $Status = "Unsupported"
                    $SettingValue += $ClientAppType + ";"
                    $Comment = $URLTeamsDevicesCA
                    $PolicyErrors++
                }
            }
            $SettingPSObj = [PSCustomObject][Ordered]@{
                ID                    = $ID
                PolicyName            = $ConditionalAccessPolicy.displayName
                PolicyID              = $ConditionalAccessPolicy.id
                PolicyState           = $PolicyState
                PolicyAccessType      = $AccessType
                AssignedToGroup       = $outAssignedToGroup
                AssignedToGroupList   = $AssignedToGroup
                ExcludedFromGroup     = $outExcludedFromGroup 
                ExcludedFromGroupList = $ExcludedFromGroup
                TeamsDevicesStatus    = $Status 
                Setting               = $Setting 
                SettingDescription    = $SettingDescription 
                Value                 = $SettingValue
                Comment               = $Comment
            }
            [void]$output.Add($SettingPSObj)
            #endregion
            
            #region 8: Conditions > Filter for devices
            $ID = 8
            $Setting = "deviceFilter"
            $SettingDescription = "Conditions > Filter for devices"
            $Comment = ""
            $DeviceFilters = ""
            if ($ConditionalAccessPolicy.conditions.devices.deviceFilter.mode -eq "exclude") {
                $Status = "Supported"
                $SettingValue = $ConditionalAccessPolicy.conditions.devices.deviceFilter.mode + ": " + $ConditionalAccessPolicy.conditions.devices.deviceFilter.rule
                $DeviceFilters = $ConditionalAccessPolicy.conditions.devices.deviceFilter.rule
            }
            else {
                $SettingValue = "Not Configured"
                $Status = "Warning"
                $Comment = "https://learn.microsoft.com/microsoftteams/troubleshoot/teams-rooms-and-devices/teams-android-devices-conditional-access-issues"
            }
            $SettingPSObj = [PSCustomObject][Ordered]@{
                ID                    = $ID
                PolicyName            = $ConditionalAccessPolicy.displayName
                PolicyID              = $ConditionalAccessPolicy.id
                PolicyState           = $PolicyState
                PolicyAccessType      = $AccessType
                AssignedToGroup       = $outAssignedToGroup
                AssignedToGroupList   = $AssignedToGroup
                ExcludedFromGroup     = $outExcludedFromGroup 
                ExcludedFromGroupList = $ExcludedFromGroup
                TeamsDevicesStatus    = $Status 
                Setting               = $Setting 
                SettingDescription    = $SettingDescription 
                Value                 = $SettingValue
                Comment               = $Comment
            }
            [void]$output.Add($SettingPSObj)
            #endregion
            
            #2024-09-24: Added check authentication flows
            #region 9: Conditions > Authentication flows
            $ID = 9
            $Setting = "authenticationFlows"
            $SettingDescription = "Conditions > Authentication flows"
            $SettingValue = "Not Configured"
            $Comment = "" 
            $Status = "Supported"
            if ($ConditionalAccessPolicy.conditions.authenticationFlows.transferMethods -like "*deviceCodeFlow*") {
                $SettingValue = $ConditionalAccessPolicy.conditions.authenticationFlows.transferMethods
                if ($isBlockAccessCA -and $DeviceType -ne "MTRWindows") {
                    $Status = "Not Supported"
                    $Comment = "Authentication flows with block will prevent Teams Devices Remote Sign in" 
                    $PolicyErrors++
                }
                else {
                    $Status = "Supported"
                }
            }
            $SettingPSObj = [PSCustomObject][Ordered]@{
                ID                    = $ID
                PolicyName            = $ConditionalAccessPolicy.displayName
                PolicyID              = $ConditionalAccessPolicy.id
                PolicyState           = $PolicyState
                PolicyAccessType      = $AccessType
                AssignedToGroup       = $outAssignedToGroup
                AssignedToGroupList   = $AssignedToGroup
                ExcludedFromGroup     = $outExcludedFromGroup 
                ExcludedFromGroupList = $ExcludedFromGroup
                TeamsDevicesStatus    = $Status 
                Setting               = $Setting 
                SettingDescription    = $SettingDescription 
                Value                 = $SettingValue
                Comment               = $Comment
            }
            [void]$output.Add($SettingPSObj)
            #endregion                    
            
            #region 10: Access controls > Grant
            $Setting = "GrantControls"
            foreach ($BuiltInControl in $ConditionalAccessPolicy.GrantControls.BuiltInControls) {
                $Comment = "" 
                $SettingValue = "Enabled"
                switch ($BuiltInControl) {
                    "mfa" {
                        $ID = 11
                        $Status = "Warning"
                        $SettingDescription = "Access controls > Grant > Require multi-factor authentication"
                        $PolicyWarnings++
                        $Comment = "If user-interactive MFA is enforced with conditional access policies. You may use per-user MFA to unblock DCF sign-in temporarily on but this is deprecated in September 2025. https://learn.microsoft.com/en-us/MicrosoftTeams/rooms/android-migration-guide#step-3---considerations-before-deploying-aosp-dm-capable-migration-firmware" 
                        if ($hasIntune) {
                            $Comment = "Require MFA will likely to cause problems during/after AOSP migration." 
                        }
                        if ($DeviceType -eq "MTRWindows") {
                            $Status = "Unsupported"
                            $Comment = "Require multi-factor authentication not supported for MTR Windows."
                            $PolicyErrors++
                        }
                    }
                    "compliantDevice" {
                        $ID = 12
                        $Status = "Supported"
                        $SettingDescription = "Access controls > Grant > Require device to be marked as compliant"
                    }
                    "DomainJoinedDevice" { 
                        $ID = 13
                        $Status = "Unsupported"
                        $SettingDescription = "Access controls > Grant > Require Hybrid Azure AD joined device"
                        $PolicyErrors++
                    }
                    "ApprovedApplication" { 
                        $ID = 14
                        $Status = "Unsupported"
                        $SettingDescription = "Access controls > Grant > Require approved client app"
                        $Comment = $URLTeamsDevicesCA
                        $PolicyErrors++
                    }
                    "CompliantApplication" { 
                        $ID = 15
                        $Status = "Unsupported"
                        $SettingDescription = "Access controls > Grant > Require app protection policy"
                        $Comment = $URLTeamsDevicesCA
                        $PolicyErrors++
                    }
                    "PasswordChange" { 
                        $ID = 16
                        $Status = "Unsupported"
                        $SettingDescription = "Access controls > Grant > Require password change"
                        $Comment = $URLTeamsDevicesCA 
                        $PolicyErrors++
                    }
                    default { 
                        $ID = 10
                        $SettingDescription = "Access controls > Grant > " + $BuiltInControl
                        $Status = "Supported"
                    }
                }
                $SettingPSObj = [PSCustomObject][Ordered]@{
                    ID                    = $ID
                    PolicyName            = $ConditionalAccessPolicy.displayName
                    PolicyID              = $ConditionalAccessPolicy.id
                    PolicyState           = $PolicyState
                    PolicyAccessType      = $AccessType
                    AssignedToGroup       = $outAssignedToGroup
                    AssignedToGroupList   = $AssignedToGroup
                    ExcludedFromGroup     = $outExcludedFromGroup 
                    ExcludedFromGroupList = $ExcludedFromGroup
                    TeamsDevicesStatus    = $Status 
                    Setting               = $Setting 
                    SettingDescription    = $SettingDescription 
                    Value                 = $SettingValue
                    Comment               = $Comment
                }
                [void]$output.Add($SettingPSObj) 
            }
            #endregion
            
            #region 11: Multifactor Authentication in Require Authentication Strength
            if ($ConditionalAccessPolicy.GrantControls.authenticationStrength) {
                $ID = 11
                $Setting = "AuthenticationStrength"
                $SettingDescription = "Access controls > Grant > Require Authentication Strength"
                $SettingValue = "Enabled"
                $Comment = "Require authentication strength is not supported." 
                $Status = "Unsupported"
                $PolicyErrors++         
                $SettingPSObj = [PSCustomObject][Ordered]@{
                    ID                    = $ID
                    PolicyName            = $ConditionalAccessPolicy.displayName
                    PolicyID              = $ConditionalAccessPolicy.id
                    PolicyState           = $PolicyState
                    PolicyAccessType      = $AccessType
                    AssignedToGroup       = $outAssignedToGroup
                    AssignedToGroupList   = $AssignedToGroup
                    ExcludedFromGroup     = $outExcludedFromGroup 
                    ExcludedFromGroupList = $ExcludedFromGroup
                    TeamsDevicesStatus    = $Status 
                    Setting               = $Setting 
                    SettingDescription    = $SettingDescription 
                    Value                 = $SettingValue
                    Comment               = $Comment
                }
                [void]$output.Add($SettingPSObj) 
            }
            #endregion
                            
            #region 17: Access controls > Grant > Custom Authentication Factors
            $ID = 17
            $Setting = "CustomAuthenticationFactors"
            $SettingDescription = "Access controls > Grant > Custom Authentication Factors"
            $SettingValue = "Not Configured"
            $Comment = "" 
            $Status = "Supported"
            if ($ConditionalAccessPolicy.GrantControls.CustomAuthenticationFactors) {
                $Status = "Unsupported"
                $SettingValue = "Enabled"
                $PolicyErrors++
                $Comment = $URLTeamsDevicesCA
            }
            $SettingPSObj = [PSCustomObject][Ordered]@{
                ID                    = $ID
                PolicyName            = $ConditionalAccessPolicy.displayName
                PolicyID              = $ConditionalAccessPolicy.id
                PolicyState           = $PolicyState
                PolicyAccessType      = $AccessType
                AssignedToGroup       = $outAssignedToGroup
                AssignedToGroupList   = $AssignedToGroup
                ExcludedFromGroup     = $outExcludedFromGroup 
                ExcludedFromGroupList = $ExcludedFromGroup
                TeamsDevicesStatus    = $Status 
                Setting               = $Setting 
                SettingDescription    = $SettingDescription 
                Value                 = $SettingValue
                Comment               = $Comment
            }
            [void]$output.Add($SettingPSObj)
            #endregion
            
            #region 18: Access controls > Grant > Terms of Use
            $ID = 18
            $Setting = "TermsOfUse"
            $SettingDescription = "Access controls > Grant > Terms of Use"
            $SettingValue = "Not Configured"
            $Comment = "" 
            $Status = "Supported"
            if ($ConditionalAccessPolicy.GrantControls.TermsOfUse) {
                $Status = "Unsupported"
                $SettingValue = "Enabled"
                $Comment = $URLTeamsDevicesKnownIssues
                $PolicyErrors++
            }
            $SettingPSObj = [PSCustomObject][Ordered]@{
                ID                    = $ID
                PolicyName            = $ConditionalAccessPolicy.displayName
                PolicyID              = $ConditionalAccessPolicy.id
                PolicyState           = $PolicyState
                PolicyAccessType      = $AccessType
                AssignedToGroup       = $outAssignedToGroup
                AssignedToGroupList   = $AssignedToGroup
                ExcludedFromGroup     = $outExcludedFromGroup 
                ExcludedFromGroupList = $ExcludedFromGroup
                TeamsDevicesStatus    = $Status 
                Setting               = $Setting 
                SettingDescription    = $SettingDescription 
                Value                 = $SettingValue
                Comment               = $Comment
            }
            [void]$output.Add($SettingPSObj)
            #endregion
            
            #region 19: Access controls > Session > Use app enforced restrictions
            $ID = 19
            $Setting = "ApplicationEnforcedRestrictions"
            $SettingDescription = "Access controls > Session > Use app enforced restrictions"
            $SettingValue = "Not Configured"
            $Comment = "" 
            $Status = "Supported"
            if ($ConditionalAccessPolicy.SessionControls.ApplicationEnforcedRestrictions) {
                $Status = "Unsupported"
                $SettingValue = "Enabled"
                $Comment = $URLTeamsDevicesCA
                $PolicyErrors++
            }
            $SettingPSObj = [PSCustomObject][Ordered]@{
                ID                    = $ID
                PolicyName            = $ConditionalAccessPolicy.displayName
                PolicyID              = $ConditionalAccessPolicy.id
                PolicyState           = $PolicyState
                PolicyAccessType      = $AccessType
                AssignedToGroup       = $outAssignedToGroup
                AssignedToGroupList   = $AssignedToGroup
                ExcludedFromGroup     = $outExcludedFromGroup 
                ExcludedFromGroupList = $ExcludedFromGroup
                TeamsDevicesStatus    = $Status 
                Setting               = $Setting 
                SettingDescription    = $SettingDescription 
                Value                 = $SettingValue
                Comment               = $Comment
            }
            [void]$output.Add($SettingPSObj)
            #endregion
                            
            #region 20: Access controls > Session > Use Conditional Access App Control
            $ID = 20
            $Setting = "CloudAppSecurity"
            $SettingDescription = "Access controls > Session > Use Conditional Access App Control"
            $SettingValue = "Not Configured"
            $Comment = "" 
            $Status = "Supported"
            if ($ConditionalAccessPolicy.SessionControls.CloudAppSecurity) {
                $Status = "Unsupported"
                $SettingValue = $ConditionalAccessPolicy.SessionControls.CloudAppSecurity.cloudAppSecurityType
                $Comment = $URLTeamsDevicesCA
                $PolicyErrors++
            }
            $SettingPSObj = [PSCustomObject][Ordered]@{
                ID                    = $ID
                PolicyName            = $ConditionalAccessPolicy.displayName
                PolicyID              = $ConditionalAccessPolicy.id
                PolicyState           = $PolicyState
                PolicyAccessType      = $AccessType
                AssignedToGroup       = $outAssignedToGroup
                AssignedToGroupList   = $AssignedToGroup
                ExcludedFromGroup     = $outExcludedFromGroup 
                ExcludedFromGroupList = $ExcludedFromGroup
                TeamsDevicesStatus    = $Status 
                Setting               = $Setting 
                SettingDescription    = $SettingDescription 
                Value                 = $SettingValue
                Comment               = $Comment
            }
            [void]$output.Add($SettingPSObj)
            #endregion
            
            #region 21: Access controls > Session > Sign-in frequency
            $ID = 21
            $Setting = "SignInFrequency"
            $SettingDescription = "Access controls > Session > Sign-in frequency"
            $SettingValue = "Not Configured"
            $Comment = "" 
            $Status = "Supported"
            if ($ConditionalAccessPolicy.SessionControls.SignInFrequency.isEnabled -eq "true") {
                if ($ConditionalAccessPolicy.SessionControls.SignInFrequency.signInFrequencyInterval -eq "everyTime") {
                    $Status = "Unsupported"
                    $Comment = "Sign In frequency set to every time will cause signin loops. https://learn.microsoft.com/en-us/entra/identity/conditional-access/concept-session-lifetime#require-reauthentication-every-time"
                    $PolicyErrors++
                }
                else {
                    $Status = "Warning"
                    $Comment = "Users will be signout from Teams Device every " + $ConditionalAccessPolicy.SessionControls.SignInFrequency.Value + " " + $ConditionalAccessPolicy.SessionControls.SignInFrequency.Type
                    $PolicyWarnings++
                }
                 $SettingValue = "" + $ConditionalAccessPolicy.SessionControls.SignInFrequency.Value + " " + $ConditionalAccessPolicy.SessionControls.SignInFrequency.Type

            }
            $SettingPSObj = [PSCustomObject][Ordered]@{
                ID                    = $ID
                PolicyName            = $ConditionalAccessPolicy.displayName
                PolicyID              = $ConditionalAccessPolicy.id
                PolicyState           = $PolicyState
                PolicyAccessType      = $AccessType
                AssignedToGroup       = $outAssignedToGroup
                AssignedToGroupList   = $AssignedToGroup
                ExcludedFromGroup     = $outExcludedFromGroup 
                ExcludedFromGroupList = $ExcludedFromGroup
                TeamsDevicesStatus    = $Status 
                Setting               = $Setting 
                SettingDescription    = $SettingDescription 
                Value                 = $SettingValue
                Comment               = $Comment
            }
            [void]$output.Add($SettingPSObj)
            #endregion
            
            #region 22: Access controls > Session > Persistent browser session
            $ID = 22
            $Setting = "PersistentBrowser"
            $SettingDescription = "Access controls > Session > Persistent browser session"
            $SettingValue = "Not Configured"
            $Comment = "" 
            $Status = "Supported"
            if ($ConditionalAccessPolicy.SessionControls.PersistentBrowser.isEnabled -eq "true") {
                $Status = "Unsupported"
                $SettingValue = $ConditionalAccessPolicy.SessionControls.persistentBrowser.mode
                $Comment = $URLTeamsDevicesCA
                $PolicyErrors++
            }
            $SettingPSObj = [PSCustomObject][Ordered]@{
                ID                    = $ID
                PolicyName            = $ConditionalAccessPolicy.displayName
                PolicyID              = $ConditionalAccessPolicy.id
                PolicyState           = $PolicyState
                PolicyAccessType      = $AccessType
                AssignedToGroup       = $outAssignedToGroup
                AssignedToGroupList   = $AssignedToGroup
                ExcludedFromGroup     = $outExcludedFromGroup 
                ExcludedFromGroupList = $ExcludedFromGroup
                TeamsDevicesStatus    = $Status 
                Setting               = $Setting 
                SettingDescription    = $SettingDescription 
                Value                 = $SettingValue
                Comment               = $Comment
            }
            [void]$output.Add($SettingPSObj)
            #endregion
            
            #region 23: Access controls > Session > Customize continuous access evaluation (CAE)
            $ID = 23
            $Setting = "ContinuousAccessEvaluation"
            $SettingDescription = "Access controls > Session > Customize continuous access evaluation (CAE)"
            $SettingValue = "Not Configured"
            $Comment = "" 
            $Status = "Supported"
            #Disable wont affected the devices and we will output in IncludedSupported.
            if ($ConditionalAccessPolicy.SessionControls.continuousAccessEvaluation.mode) {
                $SettingValue = $ConditionalAccessPolicy.SessionControls.continuousAccessEvaluation.mode
            
                if ($ConditionalAccessPolicy.SessionControls.continuousAccessEvaluation.mode -ne "disabled") {
                    $Status = "Unsupported"
                    $Comment = $URLTeamsDevicesCA
                    $PolicyErrors++
                }
            }
            $SettingPSObj = [PSCustomObject][Ordered]@{
                ID                    = $ID
                PolicyName            = $ConditionalAccessPolicy.displayName
                PolicyID              = $ConditionalAccessPolicy.id
                PolicyState           = $PolicyState
                PolicyAccessType      = $AccessType
                AssignedToGroup       = $outAssignedToGroup
                AssignedToGroupList   = $AssignedToGroup
                ExcludedFromGroup     = $outExcludedFromGroup 
                ExcludedFromGroupList = $ExcludedFromGroup
                TeamsDevicesStatus    = $Status 
                Setting               = $Setting 
                SettingDescription    = $SettingDescription 
                Value                 = $SettingValue
                Comment               = $Comment
            }
            [void]$output.Add($SettingPSObj)
            #endregion

            #region 24: Access controls > Session > Disable resilience defaults
            #endregion
            $ID = 24
            $Setting = "DisableResilienceDefaults"
            $SettingDescription = "Access controls > Session > Disable resilience defaults"
            $SettingValue = "Not Configured"
            $Comment = "" 
            $Status = "Supported"
            if ($ConditionalAccessPolicy.SessionControls.disableResilienceDefaults -eq "true") {
                $Status = "Unsupported"
                $SettingValue = $ConditionalAccessPolicy.SessionControls.disableResilienceDefaults
                $Comment = $URLTeamsDevicesCA
                $PolicyErrors++
            }
            $SettingPSObj = [PSCustomObject][Ordered]@{
                ID                    = $ID
                PolicyName            = $ConditionalAccessPolicy.displayName
                PolicyID              = $ConditionalAccessPolicy.id
                PolicyState           = $PolicyState
                PolicyAccessType      = $AccessType
                AssignedToGroup       = $outAssignedToGroup
                AssignedToGroupList   = $AssignedToGroup
                ExcludedFromGroup     = $outExcludedFromGroup 
                ExcludedFromGroupList = $ExcludedFromGroup
                TeamsDevicesStatus    = $Status 
                Setting               = $Setting 
                SettingDescription    = $SettingDescription 
                Value                 = $SettingValue
                Comment               = $Comment
            }
            [void]$output.Add($SettingPSObj)

            #region 25: Access controls > Session > Require token protection for sign-in sessions (Preview)
            $ID = 25
            $Setting = "SecureSignInSession"
            $SettingDescription = "Access controls > Session > Require token protection for sign-in sessions (Preview)"
            $SettingValue = "Not Configured"
            $Comment = "" 
            $Status = "Supported"
            if ($ConditionalAccessPolicy.SessionControls.secureSignInSession) {
                $Status = "Unsupported"
                $SettingValue = "Enabled"
                $Comment = $URLTeamsDevicesCA
                $PolicyErrors++
            }
            $SettingPSObj = [PSCustomObject][Ordered]@{
                ID                    = $ID
                PolicyName            = $ConditionalAccessPolicy.displayName
                PolicyID              = $ConditionalAccessPolicy.id
                PolicyState           = $PolicyState
                PolicyAccessType      = $AccessType
                AssignedToGroup       = $outAssignedToGroup
                AssignedToGroupList   = $AssignedToGroup
                ExcludedFromGroup     = $outExcludedFromGroup 
                ExcludedFromGroupList = $ExcludedFromGroup
                TeamsDevicesStatus    = $Status 
                Setting               = $Setting 
                SettingDescription    = $SettingDescription 
                Value                 = $SettingValue
                Comment               = $Comment
            }
            [void]$output.Add($SettingPSObj)
            #endregion

            #2025-03-11: If a policy has unsupported settings but has device filters we should ignore it.
            if (!([string]::IsNullOrEmpty($DeviceFilters)) -and ($PolicyErrors -gt 0 -or $PolicyWarnings -gt 0)) {
                $StatusSum = "Excluded with Device Filter: " + $DeviceFilters.Replace("device.", "")
            }
            elseif ($PolicyErrors -gt 0) {
                $StatusSum = "Has " + $PolicyErrors + " unsupported settings."
                $displayWarning = $true
            }
            elseif ($PolicyWarnings -gt 0) {
                $StatusSum = "Has " + $PolicyWarnings + " settings that may impact users."
                $displayWarning = $true
            }
            else {
                $StatusSum = "All settings supported."
            }
            $PolicySum = [PSCustomObject][Ordered]@{
                PolicyName            = $ConditionalAccessPolicy.DisplayName
                PolicyID              = $ConditionalAccessPolicy.id
                PolicyState           = $PolicyState
                AccessType            = $AccessType
                AssignedToGroup       = $outAssignedToGroup
                AssignedToGroupList   = $AssignedToGroup
                ExcludedFromGroup     = $outExcludedFromGroup 
                ExcludedFromGroupList = $ExcludedFromGroup
                TeamsDevicesStatus    = $StatusSum
            }
            $PolicySum.PSObject.TypeNames.Insert(0, 'TeamsDeviceConditionalAccessPolicy')
            [void]$outputSum.Add($PolicySum)
        }
        else {
            Write-Verbose ("Skipping Conditional Access Policy: " + $ConditionalAccessPolicy.DisplayName) 
            $skippedCAPolicies++
        }
    }
    #region Output
    #2025-03-28: Improving the output code readability 
    if ($totalCAPolicies -eq 0) {
        if ($UserUPN) {
            Write-Warning ("The user " + $UserUPN + " doesn't have any Compliance Policies assigned.")
        }
        else {
            Write-Warning "No Conditional Access Policies assigned to All Users, All Devices or group found. Please use Test-UcTeamsDevicesConditionalAccessPolicy -All to check all policies."
        }
        return
    }
    if (!$IncludeSupported) {
        $output = $output | Where-Object -Property TeamsDevicesStatus -NE -Value "Supported"
    }
    if ($Detailed) {
        if ($output.count -eq 0 -and !$IncludeSupported) {
            Write-Warning "No unsupported settings found, please use Test-UcTeamsDevicesConditionalAccessPolicy -IncludeSupported, to output all settings."
            return
        }
        if ($ExportCSV) {
            $output | Sort-Object PolicyName, ID | Select-Object -ExcludeProperty ID | Export-Csv -path $OutputFullPath -NoTypeInformation
            Write-Host ("Results available in: " + $OutputFullPath) -ForegroundColor Cyan
            return
        }
        else {
            return $output | Sort-Object PolicyName, ID | Select-Object -ExcludeProperty ID
        }
    }
    else {
        if (($skippedCAPolicies -gt 0) -and !$All) {
            Write-Warning ("Skipped $skippedCAPolicies conditional access policies since they will not be applied to Teams Devices (Disabled Policy/Not assigned to Group or User/Without related Target Resources)")
            Write-Warning ("Please use the All switch to check all policies: Test-UcTeamsDevicesConditionalAccessPolicy -All")
        }
        if ($displayWarning) {
            Write-Warning "One or more policies contain unsupported settings, please use Test-UcTeamsDevicesConditionalAccessPolicy -Detailed to identify the unsupported settings."
        }
        return $outputSum | Sort-Object PolicyName            
    }
    #endregion
}