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

.PARAMETER All
Will check all Conditional Access policies independently if they are assigned to a Group(s) or to Teams 

.PARAMETER IncludeSupported
Displays results for all settings in each  Conditional Access Policy

.PARAMETER UserUPN
Specifies a UserUPN that we want to check for applied Conditional Access policies

.PARAMETER ExportCSV
When present will export the detailed results to a CSV file. By defautl will save the file under the current user downloads, unless we specify the OutputPath.

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

Function Test-UcTeamsDevicesConditionalAccessPolicy {
    Param(
        [switch]$Detailed,
        [switch]$All,
        [switch]$IncludeSupported,
        [string]$UserUPN,
        [switch]$ExportCSV,
        [string]$OutputPath
    )

    $GraphURI_Users = "https://graph.microsoft.com/v1.0/users"
    $GraphURI_Groups = "https://graph.microsoft.com/v1.0/groups"
    $GraphURI_ConditionalAccess = "https://graph.microsoft.com/beta/identity/conditionalAccess/policies"

    $connectedMSGraph = $false
    $ConditionalAccessPolicies = $null
    $totalCAPolicies = 0
    $skippedCAPolicies = 0

    $URLTeamsDevicesCA = "https://aka.ms/TeamsDevicePolicies#supported-conditional-access-policies"
    $URLTeamsDevicesKnownIssues = "https://docs.microsoft.com/microsoftteams/troubleshoot/teams-rooms-and-devices/rooms-known-issues#teams-phone-devices"

    if (Test-UcMgGraphConnection -Scopes "Policy.Read.All", "Directory.Read.All") {
        Test-UcModuleUpdateAvailable -ModuleName UcLobbyTeams
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
        try {
            Write-Progress -Activity "Test-UcTeamsDevicesConditionalAccessPolicy" -Status "Getting Conditional Access Policies"
            $ConditionalAccessPolicies = (Invoke-MgGraphRequest -Uri ($GraphURI_ConditionalAccess + $GraphFilter) -Method GET).Value
            $connectedMSGraph = $true
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

        if ($connectedMSGraph) {
            $output = [System.Collections.ArrayList]::new()
            $outputSum = [System.Collections.ArrayList]::new()
            if ($UserUPN) {
                try {
                    $UserID = (Invoke-MgGraphRequest -Uri ($GraphURI_Users + "/" + $userUPN + "?`$select=id") -Method GET).id
                    $UserGroups = (Invoke-MgGraphRequest -Uri ($GraphURI_Users + "/" + $userUPN + "/transitiveMemberOf?`$select=id") -Method GET).value.id
                }
                catch [System.Net.Http.HttpRequestException] {
                    if ($PSItem.Exception.Response.StatusCode -eq "NotFound") {
                        Write-warning -Message ("User Not Found: " + $UserUPN)
                    }
                    return
                }
            }

            $Groups = New-Object 'System.Collections.Generic.Dictionary[string, string]'
            try{
            Write-Progress -Activity "Test-UcTeamsDevicesConditionalAccessPolicy" -Status "Fetching Service Principals details."
            $ServicePrincipals = Get-MgServicePrincipal -Select AppId, DisplayName -All
            } catch {}

            $p=0
            $policyCount = $ConditionalAccessPolicies.Count
            foreach ($ConditionalAccessPolicy in $ConditionalAccessPolicies) {
                $p++
                Write-Progress -Activity "Test-UcTeamsDevicesConditionalAccessPolicy" -Status ("Checking policy " +  $ConditionalAccessPolicy.displayName + " - $p of $policyCount")
                $AssignedToGroup = [System.Collections.ArrayList]::new()
                $ExcludedFromGroup = [System.Collections.ArrayList]::new()
                $AssignedToUserCount = 0
                $ExcludedFromUserCount = 0
                $outAssignedToGroup = ""
                $outExcludedFromGroup = ""
                $userExcluded = $false
                $StatusSum = ""
            
                $totalCAPolicies++
                $PolicyErrors = 0
                $PolicyWarnings = 0

                if ($UserUPN) {
                    if ($UserID -in $ConditionalAccessPolicy.conditions.users.excludeUsers) {
                        $userExcluded = $true
                        Write-Warning ("Skiping conditional access policy " + $ConditionalAccessPolicy.displayName + ", since user " + $UserUPN + " is part of Excluded Users")
                    }
                    elseif ($UserID -in $ConditionalAccessPolicy.conditions.users.includeUsers) {
                        $userIncluded = $true
                    }
                    else {
                        $userIncluded = $false
                    }
                }

                #All Users in Conditional Access Policy will show as a 'All' in the includeUsers.
                if ("All" -in $ConditionalAccessPolicy.conditions.users.includeUsers) {
                    $GroupEntry = New-Object -TypeName PSObject -Property @{
                        GroupID          = "All"
                        GroupDisplayName = "All Users"
                    }
                    $AssignedToGroup.Add($GroupEntry) | Out-Null
                    $userIncluded = $true
                }
                elseif ((($ConditionalAccessPolicy.conditions.users.includeUsers).count -gt 0) -and "None" -notin $ConditionalAccessPolicy.conditions.users.includeUsers) {
                    $AssignedToUserCount = ($ConditionalAccessPolicy.conditions.users.includeUsers).count
                    if (!$UserUPN) {
                        $userIncluded = $true
                    }
                }
                foreach ($includedGroup in $ConditionalAccessPolicy.conditions.users.includeGroups) {
                    $GroupDisplayName = $includedGroup
                    if ($Groups.ContainsKey($includedGroup)) {
                        $GroupDisplayName = $Groups.Item($includedGroup)
                    }
                    else {
                        try {
                            $GroupInfo = Invoke-MgGraphRequest -Uri ($GraphURI_Groups + "/" + $includedGroup + "/?`$select=id,displayname") -Method GET
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
                    #We only need to add if we didn't specify a UPN or if the user is part of the group that has the CA assigned.
                    if (!$UserUPN) {
                        $AssignedToGroup.Add($GroupEntry) | Out-Null
                    }
                    if ($includedGroup -in $UserGroups) {
                        $userIncluded = $true
                        $AssignedToGroup.Add($GroupEntry) | Out-Null
                    }
                }

            
                foreach ($excludedGroup in $ConditionalAccessPolicy.conditions.users.excludeGroups) {
                    $GroupDisplayName = $excludedGroup
                    if ($Groups.ContainsKey($excludedGroup)) {
                        $GroupDisplayName = $Groups.Item($excludedGroup)
                    }
                    else {
                        try {
                            $GroupInfo = Invoke-MgGraphRequest -Uri ($GraphURI_Groups + "/" + $excludedGroup + "/?`$select=id,displayname") -Method GET
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
                        Write-Warning ("Skiping conditional access policy " + $ConditionalAccessPolicy.displayName + ", since user " + $UserUPN + " is part of an Excluded Group: " + $GroupEntry.GroupDisplayName)
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
                $SettingValue = ""
                foreach ($Application in $ConditionalAccessPolicy.Conditions.Applications.IncludeApplications) {
                    $appDisplayName = ($ServicePrincipals |  Where-Object -Property AppId -eq -Value $Application).DisplayName
                    switch ($Application) {
                        "All" { $hasOffice365 = $true; $SettingValue = "All" }
                        "Office365" { $hasOffice365 = $true; $SettingValue = "Office 365" }
                        "00000002-0000-0ff1-ce00-000000000000" { $hasExchange = $true; $SettingValue += $appDisplayName + "; " }
                        "00000003-0000-0ff1-ce00-000000000000" { $hasSharePoint = $true; $SettingValue += $appDisplayName + "; " }
                        "cc15fd57-2c6c-4117-a88c-83b1d56b4bbe" { $hasTeams = $true; $SettingValue += $appDisplayName + "; " }
                        default { $SettingValue += $appDisplayName + "; " }
                    }
                }
                if ($SettingValue.EndsWith("; ")) {
                    $SettingValue = $SettingValue.Substring(0, $SettingValue.Length - 2)
                }

                if (((($AssignedToGroup.count -gt 0) -and ($hasOffice365 -or $hasTeams) -and ($PolicyState -NE "disabled")) -and (!$userExcluded) -and $userIncluded) -or $all) {

                    if (($hasExchange -and $hasSharePoint -and $hasTeams) -or ($hasOffice365)) {
                        $Status = "Supported"
                    }
                    else {
                        $Status = "Unsupported"
                        $Comment = "Teams Devices needs to access: Office 365 or Exchange Online, SharePoint Online, and Microsoft Teams"
                        $PolicyErrors++
                    }

                    $SettingPSObj = [PSCustomObject]@{
                        PolicyName            = $ConditionalAccessPolicy.displayName
                        PolicyState           = $PolicyState
                        Setting               = $Setting 
                        Value                 = $SettingValue
                        TeamsDevicesStatus    = $Status 
                        Comment               = $Comment
                        SettingDescription    = $SettingDescription 
                        AssignedToGroup       = $outAssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroupList = $ExcludedFromGroup
                        PolicyID              = $ConditionalAccessPolicy.id
                        ID                    = $ID
                    }
                    $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceConditionalAccessPolicyDetailed')
                    [void]$output.Add($SettingPSObj)
                    #endregion

                    #region 6: Assignment > Conditions > Locations
                    $ID = 6.1
                    $Setting = "includeLocations"
                    $SettingDescription = "Assignment > Conditions > Locations"
                    $Comment = ""
                    $Status = "Supported"
                    if ($ConditionalAccessPolicy.conditions.locations.includeLocations) {
                        $SettingValue = $ConditionalAccessPolicy.conditions.locations.includeLocations
                    }
                    else {
                        $SettingValue = "Not Configured"
                    }
                    $SettingPSObj = [PSCustomObject]@{
                        PolicyName            = $ConditionalAccessPolicy.displayName
                        PolicyState           = $PolicyState
                        Setting               = $Setting 
                        Value                 = $SettingValue
                        TeamsDevicesStatus    = $Status 
                        Comment               = $Comment
                        SettingDescription    = $SettingDescription 
                        AssignedToGroup       = $outAssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroupList = $ExcludedFromGroup
                        PolicyID              = $ConditionalAccessPolicy.id
                        ID                    = $ID
                    }
                    $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceConditionalAccessPolicyDetailed')
                    [void]$output.Add($SettingPSObj)

                    $ID = 6.2
                    $Setting = "excludeLocations"
                    $SettingDescription = "Assignment > Conditions > Locations"
                    $Comment = ""
                    $Status = "Supported"
                    if ($ConditionalAccessPolicy.conditions.locations.excludeLocations) {
                        $SettingValue = $ConditionalAccessPolicy.conditions.locations.excludeLocations
                    }
                    else {
                        $SettingValue = "Not Configured"
                    }
                    $SettingPSObj = [PSCustomObject]@{
                        PolicyName            = $ConditionalAccessPolicy.displayName
                        PolicyState           = $PolicyState
                        Setting               = $Setting 
                        Value                 = $SettingValue
                        TeamsDevicesStatus    = $Status 
                        Comment               = $Comment
                        SettingDescription    = $SettingDescription 
                        AssignedToGroup       = $outAssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroupList = $ExcludedFromGroup
                        PolicyID              = $ConditionalAccessPolicy.id
                        ID                    = $ID
                    }
                    $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceConditionalAccessPolicyDetailed')
                    [void]$output.Add($SettingPSObj)
                    #endregion

                    #region 7: Assignment > Conditions > Client apps
                    $ID = 7
                    $Setting = "ClientAppTypes"
                    $SettingDescription = "Assignment > Conditions > Client apps"
                    $SettingValue = ""
                    $Comment = ""
                    foreach ($ClientAppType in $ConditionalAccessPolicy.Conditions.ClientAppTypes) {
                        if ($ClientAppType -eq "All") {
                            $Status = "Supported"
                            $SettingValue = $ClientAppType
                            $Comment = ""
                        }
                        else {
                            $Status = "Unsupported"
                            $SettingValue += $ClientAppType + ";"
                            $Comment = $URLTeamsDevicesCA
                            $PolicyErrors++
                        }
                    
                    }
                    $SettingPSObj = [PSCustomObject]@{
                        PolicyName            = $ConditionalAccessPolicy.displayName
                        PolicyState           = $PolicyState
                        Setting               = $Setting 
                        Value                 = $SettingValue
                        TeamsDevicesStatus    = $Status 
                        Comment               = $Comment
                        SettingDescription    = $SettingDescription 
                        AssignedToGroup       = $outAssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroupList = $ExcludedFromGroup
                        PolicyID              = $ConditionalAccessPolicy.id
                        ID                    = $ID
                    }
                    $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceConditionalAccessPolicyDetailed')
                    [void]$output.Add($SettingPSObj)
                    #endregion

                    #region 8: Assignment > Conditions > Filter for devices
                    $ID = 8
                    $Setting = "deviceFilter"
                    $SettingDescription = "Assignment > Conditions > Filter for devices"
                    $Comment = ""
                    if ($ConditionalAccessPolicy.conditions.devices.deviceFilter.mode -eq "exclude") {
                        $Status = "Supported"
                        $SettingValue = $ConditionalAccessPolicy.conditions.devices.deviceFilter.mode + ": " + $ConditionalAccessPolicy.conditions.devices.deviceFilter.rule
                    }
                    else {
                        $SettingValue = "Not Configured"
                        $Status = "Warning"
                        $Comment = "https://learn.microsoft.com/microsoftteams/troubleshoot/teams-rooms-and-devices/teams-android-devices-conditional-access-issues"
                    }
                    $SettingPSObj = [PSCustomObject]@{
                        PolicyName            = $ConditionalAccessPolicy.displayName
                        PolicyState           = $PolicyState
                        Setting               = $Setting 
                        Value                 = $SettingValue
                        TeamsDevicesStatus    = $Status 
                        Comment               = $Comment
                        SettingDescription    = $SettingDescription 
                        AssignedToGroup       = $outAssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroupList = $ExcludedFromGroup
                        PolicyID              = $ConditionalAccessPolicy.id
                        ID                    = $ID
                    }
                    $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceConditionalAccessPolicyDetailed')
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
                                $Comment = "Require multi-factor authentication only supported for Teams Phones and Displays." 
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
                        $SettingPSObj = [PSCustomObject]@{
                            PolicyName            = $ConditionalAccessPolicy.displayName
                            PolicyState           = $PolicyState
                            Setting               = $Setting 
                            Value                 = $SettingValue
                            TeamsDevicesStatus    = $Status 
                            Comment               = $Comment
                            SettingDescription    = $SettingDescription 
                            AssignedToGroup       = $outAssignedToGroup
                            ExcludedFromGroup     = $outExcludedFromGroup 
                            AssignedToGroupList   = $AssignedToGroup
                            ExcludedFromGroupList = $ExcludedFromGroup
                            PolicyID              = $ConditionalAccessPolicy.id
                            ID                    = $ID
                        }
                        $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceConditionalAccessPolicyDetailed')
                        [void]$output.Add($SettingPSObj) 
                    }
                    #endregion
                
                    #region 17: Access controls > Grant > Custom Authentication Factors
                    $ID = 17
                    $Setting = "CustomAuthenticationFactors"
                    $SettingDescription = "Access controls > Grant > Custom Authentication Factors"
                    if ($ConditionalAccessPolicy.GrantControls.CustomAuthenticationFactors) {
                        $Status = "Unsupported"
                        $SettingValue = "Enabled"
                        $PolicyErrors++
                        $Comment = $URLTeamsDevicesCA
                    }
                    else {
                        $Status = "Supported"
                        $SettingValue = "Disabled"
                    }
                    $SettingPSObj = [PSCustomObject]@{
                        PolicyName            = $ConditionalAccessPolicy.displayName
                        PolicyState           = $PolicyState
                        Setting               = $Setting 
                        Value                 = $SettingValue
                        TeamsDevicesStatus    = $Status 
                        Comment               = $Comment
                        SettingDescription    = $SettingDescription 
                        AssignedToGroup       = $outAssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroupList = $ExcludedFromGroup
                        PolicyID              = $ConditionalAccessPolicy.id
                        ID                    = $ID
                    }
                    $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceConditionalAccessPolicyDetailed')
                    [void]$output.Add($SettingPSObj)
                    #endregion

                    #region 18: Access controls > Grant > Terms of Use
                    $ID = 18
                    $Setting = "TermsOfUse"
                    $SettingDescription = "Access controls > Grant > Terms of Use"
                    $Comment = "" 
                    if ($ConditionalAccessPolicy.GrantControls.TermsOfUse) {
                        $Status = "Warning"
                        $SettingValue = "Enabled"
                        $Comment = $URLTeamsDevicesKnownIssues
                        $PolicyWarnings++
                    }
                    else {
                        $Status = "Supported"
                        $SettingValue = "Disabled"
                    }
                    $SettingPSObj = [PSCustomObject]@{
                        PolicyName            = $ConditionalAccessPolicy.displayName
                        PolicyState           = $PolicyState
                        Setting               = $Setting 
                        Value                 = $SettingValue
                        TeamsDevicesStatus    = $Status 
                        Comment               = $Comment
                        SettingDescription    = $SettingDescription 
                        AssignedToGroup       = $outAssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroupList = $ExcludedFromGroup
                        PolicyID              = $ConditionalAccessPolicy.id
                        ID                    = $ID
                    }
                    $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceConditionalAccessPolicyDetailed')
                    [void]$output.Add($SettingPSObj)
                    #endregion

                    #region 19:  Access controls > Session > Use app enforced restrictions
                    $ID = 19
                    $Setting = "ApplicationEnforcedRestrictions"
                    $SettingDescription = "Access controls > Session > Use app enforced restrictions"
                    $Comment = "" 
                    if ($ConditionalAccessPolicy.SessionControls.ApplicationEnforcedRestrictions) {
                        $Status = "Unsupported"
                        $SettingValue = "Enabled"
                        $PolicyErrors++
                    }
                    else {
                        $Status = "Supported"
                        $SettingValue = "Disabled"
                    }
                    $SettingPSObj = [PSCustomObject]@{
                        PolicyName            = $ConditionalAccessPolicy.displayName
                        PolicyState           = $PolicyState
                        Setting               = $Setting 
                        Value                 = $SettingValue
                        TeamsDevicesStatus    = $Status 
                        Comment               = $Comment
                        SettingDescription    = $SettingDescription 
                        AssignedToGroup       = $outAssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroupList = $ExcludedFromGroup
                        PolicyID              = $ConditionalAccessPolicy.id
                        ID                    = $ID
                    }
                    $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceConditionalAccessPolicyDetailed')
                    [void]$output.Add($SettingPSObj)
                    #endregion
                
                    #region 19: Access controls > Session > Use Conditional Access App Control
                    $ID = 19
                    $Setting = "CloudAppSecurity"
                    $SettingDescription = "Access controls > Session > Use Conditional Access App Control"
                    $Comment = "" 
                    if ($ConditionalAccessPolicy.SessionControls.CloudAppSecurity) {
                        $Status = "Unsupported"
                        $SettingValue = $ConditionalAccessPolicy.SessionControls.CloudAppSecurity.cloudAppSecurityType
                        $PolicyErrors++
                    }
                    else {
                        $Status = "Supported"
                        $SettingValue = "Not Configured"
                    }
                    $SettingPSObj = [PSCustomObject]@{
                        PolicyName            = $ConditionalAccessPolicy.displayName
                        PolicyState           = $PolicyState
                        Setting               = $Setting 
                        Value                 = $SettingValue
                        TeamsDevicesStatus    = $Status 
                        Comment               = $Comment
                        SettingDescription    = $SettingDescription 
                        AssignedToGroup       = $outAssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroupList = $ExcludedFromGroup
                        PolicyID              = $ConditionalAccessPolicy.id
                        ID                    = $ID
                    }
                    $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceConditionalAccessPolicyDetailed')
                    [void]$output.Add($SettingPSObj)
                    #endregion

                    #region 20: Access controls > Session > Sign-in frequency
                    $ID = 20
                    $Setting = "SignInFrequency"
                    $SettingDescription = "Access controls > Session > Sign-in frequency"
                    $Comment = "" 
                    if ($ConditionalAccessPolicy.SessionControls.SignInFrequency.isEnabled -eq "true") {
                        $Status = "Warning"
                        $SettingValue = "" + $ConditionalAccessPolicy.SessionControls.SignInFrequency.Value + " " + $ConditionalAccessPolicy.SessionControls.SignInFrequency.Type
                        $Comment = "Users will be signout from Teams Device every " + $ConditionalAccessPolicy.SessionControls.SignInFrequency.Value + " " + $ConditionalAccessPolicy.SessionControls.SignInFrequency.Type
                        $PolicyWarnings++
                    }
                    else {
                        $Status = "Supported"
                        $SettingValue = "Not Configured"
                    }
                    $SettingPSObj = [PSCustomObject]@{
                        PolicyName            = $ConditionalAccessPolicy.displayName
                        PolicyState           = $PolicyState
                        Setting               = $Setting 
                        Value                 = $SettingValue
                        TeamsDevicesStatus    = $Status 
                        Comment               = $Comment
                        SettingDescription    = $SettingDescription 
                        AssignedToGroup       = $outAssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroupList = $ExcludedFromGroup
                        PolicyID              = $ConditionalAccessPolicy.id
                        ID                    = $ID
                    }
                    $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceConditionalAccessPolicyDetailed')
                    [void]$output.Add($SettingPSObj)
                    #endregion

                    #region 21: Access controls > Session > Persistent browser session
                    $ID = 21
                    $Setting = "PersistentBrowser"
                    $SettingDescription = "Access controls > Session > Persistent browser session"
                    $Comment = "" 
                    if ($ConditionalAccessPolicy.SessionControls.PersistentBrowser.isEnabled -eq "true") {
                        $Status = "Unsupported"
                        $SettingValue = $ConditionalAccessPolicy.SessionControls.persistentBrowser.mode
                        $PolicyErrors++
                    }
                    else {
                        $Status = "Supported"
                        $SettingValue = "Not Configured"
                    }
                
                    $SettingPSObj = [PSCustomObject]@{
                        PolicyName            = $ConditionalAccessPolicy.displayName
                        PolicyState           = $PolicyState
                        Setting               = $Setting 
                        Value                 = $SettingValue
                        TeamsDevicesStatus    = $Status 
                        Comment               = $Comment
                        SettingDescription    = $SettingDescription 
                        AssignedToGroup       = $outAssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroupList = $ExcludedFromGroup
                        PolicyID              = $ConditionalAccessPolicy.id
                        ID                    = $ID
                    }
                    $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceConditionalAccessPolicyDetailed')
                    [void]$output.Add($SettingPSObj)
                    #endregion

                    if ($PolicyErrors -gt 0) {
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
                    $PolicySum = [PSCustomObject]@{
                        PolicyID              = $ConditionalAccessPolicy.id
                        PolicyName            = $ConditionalAccessPolicy.DisplayName
                        PolicyState           = $PolicyState
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
                    $skippedCAPolicies++
                }
            }
            if ($totalCAPolicies -eq 0) {
                if ($UserUPN) {
                    Write-Warning ("The user " + $UserUPN + " doesn't have any Compliance Policies assigned.")
                }
                else {
                    Write-Warning "No Conditional Access Policies assigned to All Users, All Devices or group found. Please use Test-UcTeamsDevicesConditionalAccessPolicy -IgnoreAssigment to check all policies."
                }
            }
        
            if ($IncludeSupported -and $Detailed) {
                if ($ExportCSV) {
                    $output | Sort-Object PolicyName, ID | Select-Object PolicyName, PolicyID, PolicyState, AssignedToGroup, ExcludedFromGroup, TeamsDevicesStatus, Setting, SettingDescription, Value, Comment | Export-Csv -path $OutputFullPath -NoTypeInformation
                    Write-Host ("Results available in: " + $OutputFullPath) -ForegroundColor Cyan
                    return
                }
                else {
                    $output | Sort-Object PolicyName, ID
                }
            }
            elseif ($Detailed) {
                if ((( $output | Where-Object -Property TeamsDevicesStatus -NE -Value "Supported").count -eq 0) -and !$IncludeSupported) {
                    Write-Warning "No unsupported settings found, please use Test-UcTeamsDevicesConditionalAccessPolicy -IncludeSupported, to output all settings."
                }
                else {
                    if ($ExportCSV) {
                        $output | Where-Object -Property TeamsDevicesStatus -NE -Value "Supported" | Sort-Object PolicyName, ID | Select-Object PolicyName, PolicyID, PolicyState, AssignedToGroup, ExcludedFromGroup, TeamsDevicesStatus, Setting, SettingDescription, Value, Comment | Export-Csv -path $OutputFullPath -NoTypeInformation
                        Write-Host ("Results available in: " + $OutputFullPath) -ForegroundColor Cyan
                    }
                    else {    
                        $output | Where-Object -Property TeamsDevicesStatus -NE -Value "Supported" | Sort-Object PolicyName, ID
                    }
                }
            }
            else {
                if (($skippedCAPolicies -gt 0) -and !$All) {
                    Write-Warning ("Skipping $skippedCAPolicies conditional access policies since they will not be applied to Teams Devices")
                    Write-Warning ("Please use the All switch to check all policies: Test-UcTeamsDevicesConditionalAccessPolicy -All")
                }
                if ($displayWarning) {
                    Write-Warning "One or more policies contain unsupported settings, please use Test-UcTeamsDevicesConditionalAccessPolicy -Detailed to identify the unsupported settings."
                }
                $outputSum | Sort-Object PolicyName
            }
        }
    }
}