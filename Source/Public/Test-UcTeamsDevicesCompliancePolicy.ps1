<#
.SYNOPSIS
Validate which Intune Compliance policies are supported by Microsoft Teams Android Devices

.DESCRIPTION
This function will validate each setting in the Intune Compliance Policy to make sure they are in line with the supported settings:

    https://docs.microsoft.com/en-us/microsoftteams/rooms/supported-ca-and-compliance-policies?tabs=phones#supported-device-compliance-policies

Contributors: Traci Herr, David Paulino

Requirements: Microsoft Graph PowerShell Module (Install-Module Microsoft.Graph)

.PARAMETER Detailed
Displays test results for unsupported settings in each Intune Compliance Policy

.PARAMETER All
Will check all Intune Compliance policies independently if they are assigned to a Group(s)

.PARAMETER IncludeSupported
Displays results for all settings in each Intune Compliance Policy

.PARAMETER PolicyID
Specifies a Policy ID that will be checked if is supported by Microsoft Teams Android Devices

.PARAMETER PolicyName
Specifies a Policy Name that will be checked if is supported by Microsoft Teams Android Devices

.PARAMETER UserUPN
Specifies a UserUPN that we want to check for applied compliance policies

.PARAMETER DeviceID
Specifies DeviceID that we want to check for applied compliance policies

.EXAMPLE 
PS> Test-UcTeamsDevicesCompliancePolicy

.EXAMPLE 
PS> Test-UcTeamsDevicesCompliancePolicy -Detailed

#>
Function Test-UcTeamsDevicesCompliancePolicy {
    Param(
        [switch]$Detailed,
        [switch]$All,
        [switch]$IncludeSupported,
        [string]$PolicyID,
        [string]$PolicyName,
        [string]$UserUPN,
        [string]$DeviceID
    )

    $connectedMSGraph = $false
    $CompliancePolicies = $null
    $totalCompliancePolicies = 0
    $skippedCompliancePolicies = 0

    $GraphURI_CompliancePolicies = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies/"
    $GraphURI_Users = "https://graph.microsoft.com/v1.0/users"
    $GraphURI_Devices = "https://graph.microsoft.com/v1.0/devices"

    $SupportedAndroidCompliancePolicies = "#microsoft.graph.androidCompliancePolicy","#microsoft.graph.androidDeviceOwnerCompliancePolicy","#microsoft.graph.aospDeviceOwnerCompliancePolicy"
    $SupportedWindowsCompliancePolicies = "#microsoft.graph.windows10CompliancePolicy"

    $URLSupportedCompliancePoliciesAndroid = "https://aka.ms/TeamsDevicePolicies?tabs=phones#supported-device-compliance-policies"
    $URLSupportedCompliancePoliciesWindows = "https://aka.ms/TeamsDevicePolicies?tabs=mtr-w#supported-device-compliance-policies"

    $scopes = (Get-MgContext).Scopes

    if (!($scopes) -or !(("DeviceManagementConfiguration.Read.All" -in $scopes) -and ("Directory.Read.All" -in $scopes))) {
        Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All","Directory.Read.All"
    }

    try {
        $CompliancePolicies = (Invoke-MgGraphRequest -Uri $GraphURI_CompliancePolicies -Method GET).value
        $connectedMSGraph = $true
    }
    catch [System.Net.Http.HttpRequestException]{
        if ($PSItem.Exception.Response.StatusCode -eq "Unauthorized") {
            Write-Error "Access Denied, please make sure the user connecing to MS Graph is part of one of the following Global Reader/Intune Service Administrator/Global Administrator roles"
        }
        else {
            Write-Error $PSItem.Exception.Message
        }
    }
    catch {
        Write-Error 'Please connect to MS Graph with Connect-MgGraph -Scopes "DeviceManagementConfiguration.Read.All","Directory.Read.All" before running this script'
    }

    if ($connectedMSGraph) {
        $output = [System.Collections.ArrayList]::new()
        $outputSum = [System.Collections.ArrayList]::new()
        if($UserUPN)
        {
            try{
                $UserGroups = (Invoke-MgGraphRequest -Uri ($GraphURI_Users + "/" + $userUPN + "/transitiveMemberOf?`$select=id") -Method GET).value.id
            } catch [System.Net.Http.HttpRequestException] {
                if($PSItem.Exception.Response.StatusCode -eq "NotFound"){
                    Write-warning -Message ("User Not Found: "+ $UserUPN)
                }
                exit
            }
            #We also need to take in consideration devices that are registered to this user
            $DeviceGroups = [System.Collections.ArrayList]::new()
            $userDevices = (Invoke-MgGraphRequest -Uri ($GraphURI_Users + "/" + $userUPN + "/registeredDevices?`$select=deviceId,displayName") -Method GET).value
            foreach($userDevice in $userDevices){
                $tmpGroups = (Invoke-MgGraphRequest -Uri ($GraphURI_Devices + "(deviceId='{" + $userDevice.deviceID + "}')/transitiveMemberOf?`$select=id") -Method GET).value.id
                foreach($tmpGroup in $tmpGroups){
                    $tmpDG = New-Object -TypeName PSObject -Property @{
                        GroupId             = $tmpGroup
                        DeviceId            = $userDevice.deviceID
                        DeviceDisplayName   = $userDevice.displayName
                        }
                    $DeviceGroups.Add($tmpDG) | Out-Null
                }
            }
        }
        
        if($DeviceID){
            try{
                $DeviceGroups = (Invoke-MgGraphRequest -Uri ($GraphURI_Devices + "(deviceId='{" + $DeviceID + "}')/transitiveMemberOf?`$select=id") -Method GET).value.id
            } catch [System.Net.Http.HttpRequestException] {
                if($PSItem.Exception.Response.StatusCode -eq "BadRequest"){
                    Write-warning -Message ("Device ID Not Found: "+ $DeviceID)
                }
                exit
            }
        }

        $Groups =  Get-MgGroup -Select Id,DisplayName -All
        foreach ($CompliancePolicy in $CompliancePolicies) {
            if ((($PolicyID -eq $CompliancePolicy.id) -or ($PolicyName -eq $CompliancePolicy.displayName) -or (!$PolicyID -and !$PolicyName)) -and (($CompliancePolicy."@odata.type" -in $SupportedAndroidCompliancePolicies) -or ($CompliancePolicy."@odata.type" -in $SupportedWindowsCompliancePolicies))) {

                #We need to check if the policy has assignments (Groups/All Users/All Devices)
                $CompliancePolicyAssignments =  (Invoke-MgGraphRequest -Uri ($GraphURI_CompliancePolicies + $CompliancePolicy.id + "/assignments" ) -Method GET).value
                $AssignedToGroup = [System.Collections.ArrayList]::new()
                $ExcludedFromGroup =[System.Collections.ArrayList]::new()
                $outAssignedToGroup = ""
                $outExcludedFromGroup =""
                #We wont need to check the settings if the policy is not assigned to a user
                if($UserUPN -or $DeviceID){
                    $userOrDeviceIncluded = $false
                } else {
                    $userOrDeviceIncluded = $true
                }

                #Define the Compliance Policy type
                switch ($CompliancePolicy."@odata.type") {
                    "#microsoft.graph.androidCompliancePolicy"              { $CPType = "Android Device" }
                    "#microsoft.graph.androidDeviceOwnerCompliancePolicy"   { $CPType = "Android Enterprise" }
                    "#microsoft.graph.aospDeviceOwnerCompliancePolicy"      { $CPType = "Android (AOSP)" }
                    "#microsoft.graph.windows10CompliancePolicy"            { $CPType = "Windows 10 or later" }
                    Default { $CPType = $CompliancePolicy."@odata.type".split('.')[2] }
                }

                #Checking Compliance Policy assigments since we can skip non assigned policies.
                foreach($CompliancePolicyAssignment in $CompliancePolicyAssignments){
                    $GroupEntry = New-Object -TypeName PSObject -Property @{
                        GroupID = $CompliancePolicyAssignment.target.Groupid
                        GroupDisplayName   = ($Groups | Where-Object -Property id -EQ -Value $CompliancePolicyAssignment.target.Groupid).displayName
                    }
                    switch ($CompliancePolicyAssignment.target."@odata.type"){
                        #Policy assigned to all users
                        "#microsoft.graph.allLicensedUsersAssignmentTarget" {
                            $GroupEntry = New-Object -TypeName PSObject -Property @{
                                GroupID = "allLicensedUsersAssignment"
                                GroupDisplayName   = "All Users"
                            }
                            $AssignedToGroup.Add($GroupEntry) | Out-Null
                            $userOrDeviceIncluded = $true
                        }
                        #Policy assigned to all devices
                        "#microsoft.graph.allDevicesAssignmentTarget" {
                            $GroupEntry = New-Object -TypeName PSObject -Property @{
                                GroupID = "allDevicesAssignmentTarget"
                                GroupDisplayName   = "All Devices"
                            }
                            $AssignedToGroup.Add($GroupEntry) | Out-Null
                            $userOrDeviceIncluded = $true
                        }
                        #Group that this policy is assigned
                        "#microsoft.graph.groupAssignmentTarget" {
                            $AssignedToGroup.Add($GroupEntry) | Out-Null
                            if(($UserUPN -or $DeviceID) -and (($CompliancePolicyAssignment.target.Groupid -in $UserGroups) -or ($CompliancePolicyAssignment.target.Groupid -in $DeviceGroups))){
                                $userOrDeviceIncluded = $true
                            }
                        }
                        #Group that this policy is excluded
                        "#microsoft.graph.exclusionGroupAssignmentTarget" {
                            $ExcludedFromGroup.Add($GroupEntry)| Out-Null
                            #If user is excluded then we dont need to check the policy
                            if($UserUPN -and ($CompliancePolicyAssignment.target.Groupid -in $UserGroups)){
                                Write-Warning ("Skiping compliance policy " +$CompliancePolicy.displayName + ", since user " + $UserUPN +" is part of an Excluded Group: " + $GroupEntry.GroupDisplayName)
                                $userOrDeviceExcluded = $true
                            } elseif($DeviceID -and ($CompliancePolicyAssignment.target.Groupid -in $DeviceGroups)){
                                Write-Warning ("Skiping compliance policy " +$CompliancePolicy.displayName + ", since device " + $DeviceID +" is part of an Excluded Group: " + $GroupEntry.GroupDisplayName)
                                $userOrDeviceExcluded = $true
                            } elseif($UserUPN -and (($CompliancePolicyAssignment.target.Groupid -in $DeviceGroups.GroupId))){
                                #In case a device is excluded we will check the policy but output a message
                                $tmpDev = ($DeviceGroups | Where-Object -Property GroupId -eq -Value $CompliancePolicyAssignment.target.Groupid)
                                Write-Warning ("Compliance policy " +$CompliancePolicy.displayName + " will not be applied to device " + $tmpDev.DeviceDisplayName +" (" + $tmpDev.DeviceID +"), since this device is part of an Excluded Group: " + $GroupEntry.GroupDisplayName)
                            }
                        }
                    }
                }
                
                if((($AssignedToGroup.count -gt 0) -and !$userOrDeviceExcluded -and $userOrDeviceIncluded) -or $all){
                    $totalCompliancePolicies++ 
                    $PolicyErrors = 0
                    $PolicyWarnings = 0

                    #If only assigned/excluded from a group we will show the group display name, otherwise the number of groups assigned/excluded.
                    if($AssignedToGroup.count -eq 1){
                        $outAssignedToGroup = $AssignedToGroup.GroupDisplayName
                    } elseif($AssignedToGroup.count -eq 0){
                        $outAssignedToGroup = "None"
                    } else {
                        $outAssignedToGroup = "" + $AssignedToGroup.count + " groups"
                    }

                    if($ExcludedFromGroup.count -eq 1){
                        $outExcludedFromGroup = $ExcludedFromGroup.GroupDisplayName
                    } elseif($ExcludedFromGroup.count -eq 0){
                        $outExcludedFromGroup = "None"
                    } else {
                        $outExcludedFromGroup = "" + $ExcludedFromGroup.count + " groups"
                    }

                    if($CompliancePolicy."@odata.type" -in $SupportedAndroidCompliancePolicies){
                        $URLSupportedCompliancePolicies = $URLSupportedCompliancePoliciesAndroid
                    }
                    elseif($CompliancePolicy."@odata.type" -in $SupportedWindowsCompliancePolicies){
                        $URLSupportedCompliancePolicies = $URLSupportedCompliancePoliciesWindows
                    }

                #region Common settings between Android and Windows 
                    #region 9: Device Properties > Operation System Version
                    $ID = 9.1
                    $Setting = "osMinimumVersion"
                    $SettingDescription = "Device Properties > Operation System Version > Minimum OS version"
                    $SettingValue = "Not Configured"
                    $Comment = ""
                    if (!([string]::IsNullOrEmpty($CompliancePolicy.osMinimumVersion))) {
                        if ($CompliancePolicy."@odata.type" -in $SupportedWindowsCompliancePolicies){
                            $Status = "Unsupported"
                            $Comment = "Teams Rooms automatically updates to newer versions of Windows and setting values here could prevent successful sign-in after an OS update."
                            $PolicyErrors++
                        } else {
                            $Status = "Warning"
                            $Comment = "This setting can cause sign in issues."
                            $PolicyWarnings++
                        }
                        $SettingValue = $CompliancePolicy.osMinimumVersion
                    }
                    else {
                        $Status = "Supported"
                    }
                    $SettingPSObj =  [PSCustomObject]@{
                        PolicyName              = $CompliancePolicy.displayName
                        PolicyType              = $CPType
                        Setting                 = $Setting
                        Value                   = $SettingValue
                        TeamsDevicesStatus      = $Status 
                        Comment                 = $Comment
                        SettingDescription      = $SettingDescription
                        AssignedToGroup         = $outAssignedToGroup
                        ExcludedFromGroup       = $outExcludedFromGroup 
                        AssignedToGroupList     = $AssignedToGroup
                        ExcludedFromGroupList   = $ExcludedFromGroup
                        PolicyID                = $CompliancePolicy.id
                        ID                      = $ID
                    }
                    $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicyDetailed')
                    [void]$output.Add($SettingPSObj)
            
                    $ID = 9.2
                    $Setting = "osMaximumVersion"
                    $SettingDescription = "Device Properties > Operation System Version > Maximum OS version"
                    $SettingValue = "Not Configured"
                    $Comment = ""
                    if (!([string]::IsNullOrEmpty($CompliancePolicy.osMaximumVersion))) {
                        if ($CompliancePolicy."@odata.type" -in $SupportedWindowsCompliancePolicies){
                            $Status = "Unsupported"
                            $Comment = "Teams Rooms automatically updates to newer versions of Windows and setting values here could prevent successful sign-in after an OS update."
                            $PolicyErrors++
                        } else {
                            $Status = "Warning"
                            $Comment = "This setting can cause sign in issues."
                            $PolicyWarnings++
                        }
                        $SettingValue = $CompliancePolicy.osMaximumVersion
                    }
                    else {
                        $Status = "Supported"
                    }
                    $SettingPSObj =  [PSCustomObject]@{
                        PolicyName              = $CompliancePolicy.displayName
                        PolicyType              = $CPType
                        Setting                 = $Setting
                        Value                   = $SettingValue
                        TeamsDevicesStatus      = $Status 
                        Comment                 = $Comment
                        SettingDescription      = $SettingDescription
                        AssignedToGroup         = $outAssignedToGroup
                        ExcludedFromGroup       = $outExcludedFromGroup 
                        AssignedToGroupList     = $AssignedToGroup
                        ExcludedFromGroupList   = $ExcludedFromGroup
                        PolicyID                = $CompliancePolicy.id
                        ID                      = $ID
                    }
                    $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicyDetailed')
                    [void]$output.Add($SettingPSObj)
                    #endregion

                    #region 17: System Security > All Android devices > Require a password to unlock mobile devices
                    $ID = 17
                    $Setting = "passwordRequired"
                    $SettingDescription = "System Security > All Android devices > Require a password to unlock mobile devices"
                    $SettingValue = "Not Configured"
                    $Comment = ""
                    if ($CompliancePolicy.passwordRequired) {
                        $Status = "Unsupported"
                        $SettingValue = "Require"
                        $Comment = $URLSupportedCompliancePolicies
                        $PolicyErrors++
                    }
                    else {
                        $Status = "Supported"
                    }
                    $SettingPSObj =  [PSCustomObject]@{
                        PolicyName              = $CompliancePolicy.displayName
                        PolicyType              = $CPType
                        Setting                 = $Setting
                        Value                   = $SettingValue
                        TeamsDevicesStatus      = $Status 
                        Comment                 = $Comment
                        SettingDescription      = $SettingDescription
                        AssignedToGroup         = $outAssignedToGroup
                        ExcludedFromGroup       = $outExcludedFromGroup 
                        AssignedToGroupList     = $AssignedToGroup
                        ExcludedFromGroupList   = $ExcludedFromGroup
                        PolicyID                = $CompliancePolicy.id
                        ID                      = $ID
                    }
                    $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicyDetailed')
                    [void]$output.Add($SettingPSObj)
                    #endregion
                #endregion

                    if($CompliancePolicy."@odata.type" -in $SupportedAndroidCompliancePolicies){

                        #region 1: Microsoft Defender for Endpoint > Require the device to be at or under the machine risk score
                        $ID = 1
                        $Setting = "deviceThreatProtectionEnabled"
                        $SettingDescription = "Microsoft Defender for Endpoint > Require the device to be at or under the machine risk score"
                        $SettingValue = "Not Configured"
                        $Comment = ""
                        if ($CompliancePolicy.deviceThreatProtectionEnabled) {
                            $Status = "Unsupported"
                            $PolicyErrors++
                            $SettingValue = $CompliancePolicy.advancedThreatProtectionRequiredSecurityLevel
                            $Comment = $URLSupportedCompliancePolicies
                        }
                        else {
                            $Status = "Supported"
                        }
                        $SettingPSObj =  [PSCustomObject]@{
                            PolicyName              = $CompliancePolicy.displayName
                            PolicyType              = $CPType
                            Setting                 = $Setting
                            Value                   = $SettingValue
                            TeamsDevicesStatus      = $Status 
                            Comment                 = $Comment
                            SettingDescription      = $SettingDescription
                            AssignedToGroup         = $outAssignedToGroup
                            ExcludedFromGroup       = $outExcludedFromGroup 
                            AssignedToGroupList     = $AssignedToGroup
                            ExcludedFromGroupList   = $ExcludedFromGroup
                            PolicyID                = $CompliancePolicy.id
                            ID                      = $ID
                        }
                        $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicyDetailed')
                        [void]$output.Add($SettingPSObj)
                        #endregion

                        #region 2: Device Health > Device managed with device administrator
                        $ID = 2
                        $Setting = "securityBlockDeviceAdministratorManagedDevices"
                        $SettingDescription = "Device Health > Device managed with device administrator"
                        $SettingValue = "Not Configured"
                        $Comment = ""
                        if ($CompliancePolicy.securityBlockDeviceAdministratorManagedDevices) {
                            $Status = "Unsupported"
                            $SettingValue = "Block"
                            $Comment = 	"Teams Android devices management requires device administrator to be enabled."
                            $PolicyErrors++
                        }
                        else {
                            $Status = "Supported"
                        }
                        $SettingPSObj =  [PSCustomObject]@{
                            PolicyName              = $CompliancePolicy.displayName
                            PolicyType              = $CPType
                            Setting                 = $Setting
                            Value                   = $SettingValue
                            TeamsDevicesStatus      = $Status 
                            Comment                 = $Comment
                            SettingDescription      = $SettingDescription
                            AssignedToGroup         = $outAssignedToGroup
                            ExcludedFromGroup       = $outExcludedFromGroup 
                            AssignedToGroupList     = $AssignedToGroup
                            ExcludedFromGroupList   = $ExcludedFromGroup
                            PolicyID                = $CompliancePolicy.id
                            ID                      = $ID
                        }
                        $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicyDetailed')
                        [void]$output.Add($SettingPSObj)
                        #endregion

                        #region 3: Device Health > Rooted devices
                        $ID = 3
                        $Setting = "securityBlockJailbrokenDevices"
                        $SettingDescription = "Device Health > Rooted devices"
                        $SettingValue = "Not Configured"
                        $Comment = ""
                        if ($CompliancePolicy.securityBlockJailbrokenDevices) {
                            $Status = "Warning"
                            $SettingValue = "Block"
                            $Comment = "This setting can cause sign in issues."
                            $PolicyWarnings++
                        }
                        else {
                            $Status = "Supported"
                        }
                        $SettingPSObj =  [PSCustomObject]@{
                            PolicyName              = $CompliancePolicy.displayName
                            PolicyType              = $CPType
                            Setting                 = $Setting
                            Value                   = $SettingValue
                            TeamsDevicesStatus      = $Status 
                            Comment                 = $Comment
                            SettingDescription      = $SettingDescription
                            AssignedToGroup         = $outAssignedToGroup
                            ExcludedFromGroup       = $outExcludedFromGroup 
                            AssignedToGroupList     = $AssignedToGroup
                            ExcludedFromGroupList   = $ExcludedFromGroup
                            PolicyID                = $CompliancePolicy.id
                            ID                      = $ID
                        }
                        $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicyDetailed')
                        [void]$output.Add($SettingPSObj)
                        #endregion

                        #region 4: Device Health > Require the device to be at or under the Device Threat Level
                        $ID = 4
                        $Setting = "deviceThreatProtectionRequiredSecurityLevel"
                        $SettingDescription = "Device Health > Require the device to be at or under the Device Threat Level"
                        $SettingValue = "Not Configured"
                        $Comment = ""
                        if ($CompliancePolicy.deviceThreatProtectionRequiredSecurityLevel -ne "unavailable") {
                            $Status = "Unsupported"
                            $SettingValue = $CompliancePolicy.deviceThreatProtectionRequiredSecurityLevel
                            $Comment = $URLSupportedCompliancePolicies
                            $PolicyErrors++
                        }
                        else {
                            $Status = "Supported"
                        }
                        $SettingPSObj =  [PSCustomObject]@{
                            PolicyName              = $CompliancePolicy.displayName
                            PolicyType              = $CPType
                            Setting                 = $Setting
                            Value                   = $SettingValue
                            TeamsDevicesStatus      = $Status 
                            Comment                 = $Comment
                            SettingDescription      = $SettingDescription
                            AssignedToGroup         = $outAssignedToGroup
                            ExcludedFromGroup       = $outExcludedFromGroup 
                            AssignedToGroupList     = $AssignedToGroup
                            ExcludedFromGroupList   = $ExcludedFromGroup
                            PolicyID                = $CompliancePolicy.id
                            ID                      = $ID
                        }
                        $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicyDetailed')
                        [void]$output.Add($SettingPSObj)
                        #endregion

                        #region 5: Device Health > Google Protect > Google Play Services is Configured
                        $ID = 5
                        $Setting = "securityRequireGooglePlayServices"
                        $SettingDescription = "Device Health > Google Protect > Google Play Services is Configured"
                        $SettingValue = "Not Configured"
                        $Comment = ""
                        if ($CompliancePolicy.securityRequireGooglePlayServices) {
                            $Status = "Unsupported"
                            $SettingValue = "Require"
                            $Comment = "Google play isn't installed on Teams Android devices."
                            $PolicyErrors++
                        }
                        else {
                            $Status = "Supported"
                        }
                        $SettingPSObj =  [PSCustomObject]@{
                            PolicyName              = $CompliancePolicy.displayName
                            PolicyType              = $CPType
                            Setting                 = $Setting
                            Value                   = $SettingValue
                            TeamsDevicesStatus      = $Status 
                            Comment                 = $Comment
                            SettingDescription      = $SettingDescription
                            AssignedToGroup         = $outAssignedToGroup
                            ExcludedFromGroup       = $outExcludedFromGroup 
                            AssignedToGroupList     = $AssignedToGroup
                            ExcludedFromGroupList   = $ExcludedFromGroup
                            PolicyID                = $CompliancePolicy.id
                            ID                      = $ID
                        }
                        $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicyDetailed')
                        [void]$output.Add($SettingPSObj)
                        #endregion

                        #region 6: Device Health > Google Protect > Up-to-date security provider
                        $ID = 6
                        $Setting = "securityRequireUpToDateSecurityProviders"
                        $SettingDescription = "Device Health > Google Protect > Up-to-date security provider"
                        $SettingValue = "Not Configured"
                        $Comment = ""
                        if ($CompliancePolicy.securityRequireUpToDateSecurityProviders) {
                            $Status = "Unsupported"
                            $SettingValue = "Require"
                            $Comment = "Google play isn't installed on Teams Android devices."
                            $PolicyErrors++
                        }
                        else {
                            $Status = "Supported"
                        }
                        $SettingPSObj =  [PSCustomObject]@{
                            PolicyName              = $CompliancePolicy.displayName
                            PolicyType              = $CPType
                            Setting                 = $Setting
                            Value                   = $SettingValue
                            TeamsDevicesStatus      = $Status 
                            Comment                 = $Comment
                            SettingDescription      = $SettingDescription
                            AssignedToGroup         = $outAssignedToGroup
                            ExcludedFromGroup       = $outExcludedFromGroup 
                            AssignedToGroupList     = $AssignedToGroup
                            ExcludedFromGroupList   = $ExcludedFromGroup
                            PolicyID                = $CompliancePolicy.id
                            ID                      = $ID
                        }
                        $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicyDetailed')
                        [void]$output.Add($SettingPSObj)
                        #endregion
                
                        #region 7: Device Health > Google Protect > Threat scan on apps
                        $ID = 7
                        $Setting = "securityRequireVerifyApps"
                        $SettingDescription = "Device Health > Google Protect > Threat scan on apps"
                        $SettingValue = "Not Configured"
                        $Comment = ""
                        if ($CompliancePolicy.securityRequireVerifyApps) {
                            $Status = "Unsupported"
                            $SettingValue = "Require"
                            $Comment = "Google play isn't installed on Teams Android devices."
                            $PolicyErrors++
                        }
                        else {
                            $Status = "Supported"
                        }
                        $SettingPSObj =  [PSCustomObject]@{
                            PolicyName              = $CompliancePolicy.displayName
                            PolicyType              = $CPType
                            Setting                 = $Setting
                            Value                   = $SettingValue
                            TeamsDevicesStatus      = $Status 
                            Comment                 = $Comment
                            SettingDescription      = $SettingDescription
                            AssignedToGroup         = $outAssignedToGroup
                            ExcludedFromGroup       = $outExcludedFromGroup 
                            AssignedToGroupList     = $AssignedToGroup
                            ExcludedFromGroupList   = $ExcludedFromGroup
                            PolicyID                = $CompliancePolicy.id
                            ID                      = $ID
                        }
                        $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicyDetailed')
                        [void]$output.Add($SettingPSObj)
                        #endregion

                        #region 8: Device Health > Google Protect > SafetyNet device attestation
                        $ID = 8
                        $Setting = "securityRequireSafetyNetAttestation"
                        $SettingDescription = "Device Health > Google Protect > SafetyNet device attestation"
                        $SettingValue = "Not Configured"
                        $Comment = ""
                        if (($CompliancePolicy.securityRequireSafetyNetAttestationBasicIntegrity) -or ($CompliancePolicy.securityRequireSafetyNetAttestationCertifiedDevice)) {
                            $Status = "Unsupported"
                            $Comment = "Google play isn't installed on Teams Android devices."
                            $PolicyErrors++
                            if ($CompliancePolicy.securityRequireSafetyNetAttestationCertifiedDevice){
                                $SettingValue = "Check basic integrity and certified devices"
                            } elseif ($CompliancePolicy.securityRequireSafetyNetAttestationBasicIntegrity){
                                $SettingValue = "Check basic integrity"
                            }
                        }
                        else {
                            $Status = "Supported"
                        }
                        $SettingPSObj =  [PSCustomObject]@{
                            PolicyName              = $CompliancePolicy.displayName
                            PolicyType              = $CPType
                            Setting                 = $Setting
                            Value                   = $SettingValue
                            TeamsDevicesStatus      = $Status 
                            Comment                 = $Comment
                            SettingDescription      = $SettingDescription
                            AssignedToGroup         = $outAssignedToGroup
                            ExcludedFromGroup       = $outExcludedFromGroup 
                            AssignedToGroupList     = $AssignedToGroup
                            ExcludedFromGroupList   = $ExcludedFromGroup
                            PolicyID                = $CompliancePolicy.id
                            ID                      = $ID
                        }
                        $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicyDetailed')
                        [void]$output.Add($SettingPSObj)
                        #endregion

                        

                        #region 10: System Security > Encryption > Require encryption of data storage on device.
                        $ID = 10
                        $Setting = "storageRequireEncryption"
                        $SettingDescription = "System Security > Encryption > Require encryption of data storage on device"
                        $SettingValue = "Not Configured"
                        $Comment = ""
                        if ($CompliancePolicy.storageRequireEncryption) {
                            $Status = "Warning"
                            $SettingValue = "Require"
                            $Comment = "Manufacturers might configure encryption attributes on their devices in a way that Intune doesn't recognize. If this happens, Intune marks the device as noncompliant."
                            $PolicyWarnings++
                        }
                        else {
                            $Status = "Supported"
                        }
                        $SettingPSObj =  [PSCustomObject]@{
                            PolicyName              = $CompliancePolicy.displayName
                            PolicyType              = $CPType
                            Setting                 = $Setting
                            Value                   = $SettingValue
                            TeamsDevicesStatus      = $Status 
                            Comment                 = $Comment
                            SettingDescription      = $SettingDescription
                            AssignedToGroup         = $outAssignedToGroup
                            ExcludedFromGroup       = $outExcludedFromGroup 
                            AssignedToGroupList     = $AssignedToGroup
                            ExcludedFromGroupList   = $ExcludedFromGroup
                            PolicyID                = $CompliancePolicy.id
                            ID                      = $ID
                        }
                        $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicyDetailed')
                        [void]$output.Add($SettingPSObj)
                        #endregion

                        #region 11: System Security > Device Security > Block apps from unknown sources
                        $ID = 11
                        $Setting = "securityPreventInstallAppsFromUnknownSources"
                        $SettingDescription = "System Security > Device Security > Block apps from unknown sources"
                        $SettingValue = "Not Configured"
                        $Comment = ""
                        if ($CompliancePolicy.securityPreventInstallAppsFromUnknownSources) {
                            $Status = "Unsupported"
                            $SettingValue = "Block"
                            $Comment = "Only Teams admins install apps or OEM tools"
                            $PolicyErrors++
                        }
                        else {
                            $Status = "Supported"
                        }
                        $SettingPSObj =  [PSCustomObject]@{
                            PolicyName              = $CompliancePolicy.displayName
                            PolicyType              = $CPType
                            Setting                 = $Setting
                            Value                   = $SettingValue
                            TeamsDevicesStatus      = $Status 
                            Comment                 = $Comment
                            SettingDescription      = $SettingDescription
                            AssignedToGroup         = $outAssignedToGroup
                            ExcludedFromGroup       = $outExcludedFromGroup 
                            AssignedToGroupList     = $AssignedToGroup
                            ExcludedFromGroupList   = $ExcludedFromGroup
                            PolicyID                = $CompliancePolicy.id
                            ID                      = $ID
                        }
                        $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicyDetailed')
                        [void]$output.Add($SettingPSObj)
                        #endregion

                        #region 14: System Security > Device Security > Minimum security patch level
                        $ID = 14
                        $Setting = "minAndroidSecurityPatchLevel"
                        $SettingDescription = "System Security > Device Security > Minimum security patch level"
                        $SettingValue = "Not Configured"
                        $Comment = ""
                        if (!([string]::IsNullOrEmpty($CompliancePolicy.minAndroidSecurityPatchLevel))){
                            $Status = "Warning"
                            $SettingValue = $CompliancePolicy.minAndroidSecurityPatchLevel
                            $Comment = "This setting can cause sign in issues."
                            $PolicyWarnings++
                        }
                        else {
                            $Status = "Supported"
                        }
                        $SettingPSObj =  [PSCustomObject]@{
                            PolicyName              = $CompliancePolicy.displayName
                            PolicyType              = $CPType
                            Setting                 = $Setting
                            Value                   = $SettingValue
                            TeamsDevicesStatus      = $Status 
                            Comment                 = $Comment
                            SettingDescription      = $SettingDescription
                            AssignedToGroup         = $outAssignedToGroup
                            ExcludedFromGroup       = $outExcludedFromGroup 
                            AssignedToGroupList     = $AssignedToGroup
                            ExcludedFromGroupList   = $ExcludedFromGroup
                            PolicyID                = $CompliancePolicy.id
                            ID                      = $ID
                        }
                        $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicyDetailed')
                        [void]$output.Add($SettingPSObj)
                        #endregion

                        #region 15: System Security > Device Security > Restricted apps
                        $ID = 15
                        $Setting = "securityPreventInstallAppsFromUnknownSources"
                        $SettingDescription = "System Security > Device Security > Restricted apps"
                        $SettingValue = "Not Configured"
                        $Comment = ""
                        if (($CompliancePolicy.restrictedApps).count -gt 0 ) {
                            $Status = "Unsupported"
                            $SettingValue = "Found " + ($CompliancePolicy.restrictedApps).count  + " restricted app(s)"
                            $Comment = $URLSupportedCompliancePolicies
                            $PolicyErrors++
                        }
                        else {
                            $Status = "Supported"
                        }
                        $SettingPSObj =  [PSCustomObject]@{
                            PolicyName              = $CompliancePolicy.displayName
                            PolicyType              = $CPType
                            Setting                 = $Setting
                            Value                   = $SettingValue
                            TeamsDevicesStatus      = $Status 
                            Comment                 = $Comment
                            SettingDescription      = $SettingDescription
                            AssignedToGroup         = $outAssignedToGroup
                            ExcludedFromGroup       = $outExcludedFromGroup 
                            AssignedToGroupList     = $AssignedToGroup
                            ExcludedFromGroupList   = $ExcludedFromGroup
                            PolicyID                = $CompliancePolicy.id
                            ID                      = $ID
                        }
                        $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicyDetailed')
                        [void]$output.Add($SettingPSObj)
                        #endregion

                        #region 16: System Security > All Android devices > Maximum minutes of inactivity before password is required
                        $ID = 16
                        $Setting = "passwordMinutesOfInactivityBeforeLock"
                        $SettingDescription = "System Security > All Android devices > Maximum minutes of inactivity before password is required"
                        $SettingValue = "Not Configured"
                        $Comment = ""
                        if (!([string]::IsNullOrEmpty($CompliancePolicy.passwordMinutesOfInactivityBeforeLock))) {
                            $Status = "Unsupported"
                            $SettingValue = "" + $CompliancePolicy.passwordMinutesOfInactivityBeforeLock + " minutes"
                            $Comment = $URLSupportedCompliancePolicies
                            $PolicyErrors++
                        }
                        else {
                            $Status = "Supported"
                        }
                        $SettingPSObj =  [PSCustomObject]@{
                            PolicyName              = $CompliancePolicy.displayName
                            PolicyType              = $CPType
                            Setting                 = $Setting
                            Value                   = $SettingValue
                            TeamsDevicesStatus      = $Status 
                            Comment                 = $Comment
                            SettingDescription      = $SettingDescription
                            AssignedToGroup         = $outAssignedToGroup
                            ExcludedFromGroup       = $outExcludedFromGroup 
                            AssignedToGroupList     = $AssignedToGroup
                            ExcludedFromGroupList   = $ExcludedFromGroup
                            PolicyID                = $CompliancePolicy.id
                            ID                      = $ID
                        }
                        $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicyDetailed')
                        [void]$output.Add($SettingPSObj)
                        #endregion
                    } 
                    elseif($CompliancePolicy."@odata.type" -in $SupportedWindowsCompliancePolicies){

                        #region 18: Device Properties > Operation System Version
                        $ID = 18.1
                        $Setting = "mobileOsMinimumVersion"
                        $SettingDescription = "Device Properties > Operation System Version > Minimum OS version for mobile devices"
                        $SettingValue = "Not Configured"
                        $Comment = ""
                        if (!([string]::IsNullOrEmpty($CompliancePolicy.mobileOsMinimumVersion))) {
                            $Status = "Unsupported"
                            $SettingValue = $CompliancePolicy.mobileOsMinimumVersion
                            $Comment = $URLSupportedCompliancePolicies
                            $PolicyErrors++
                        }
                        else {
                            $Status = "Supported"
                        }
                        $SettingPSObj =  [PSCustomObject]@{
                            PolicyName              = $CompliancePolicy.displayName
                            PolicyType              = $CPType
                            Setting                 = $Setting
                            Value                   = $SettingValue
                            TeamsDevicesStatus      = $Status 
                            Comment                 = $Comment
                            SettingDescription      = $SettingDescription
                            AssignedToGroup         = $outAssignedToGroup
                            ExcludedFromGroup       = $outExcludedFromGroup 
                            AssignedToGroupList     = $AssignedToGroup
                            ExcludedFromGroupList   = $ExcludedFromGroup
                            PolicyID                = $CompliancePolicy.id
                            ID                      = $ID
                        }
                        $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicyDetailed')
                        [void]$output.Add($SettingPSObj)
                
                        $ID = 18.2
                        $Setting = "mobileOsMaximumVersion"
                        $SettingDescription = "Device Properties > Operation System Version > Maximum OS version for mobile devices"
                        $SettingValue = "Not Configured"
                        $Comment = ""
                        if (!([string]::IsNullOrEmpty($CompliancePolicy.mobileOsMaximumVersion))) {
                            $Status = "Unsupported"
                            $SettingValue = $CompliancePolicy.mobileOsMaximumVersion
                            $Comment = $URLSupportedCompliancePolicies
                            $PolicyErrors++
                        }
                        else {
                            $Status = "Supported"
                        }
                        $SettingPSObj =  [PSCustomObject]@{
                            PolicyName              = $CompliancePolicy.displayName
                            PolicyType              = $CPType
                            Setting                 = $Setting
                            Value                   = $SettingValue
                            TeamsDevicesStatus      = $Status 
                            Comment                 = $Comment
                            SettingDescription      = $SettingDescription
                            AssignedToGroup         = $outAssignedToGroup
                            ExcludedFromGroup       = $outExcludedFromGroup 
                            AssignedToGroupList     = $AssignedToGroup
                            ExcludedFromGroupList   = $ExcludedFromGroup
                            PolicyID                = $CompliancePolicy.id
                            ID                      = $ID
                        }
                        $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicyDetailed')
                        [void]$output.Add($SettingPSObj)
                        #endregion

                        #region 19: Device Properties > Operation System Version > Valid operating system builds
                        $ID = 19
                        $Setting = "validOperatingSystemBuildRanges"
                        $SettingDescription = "Device Properties > Operation System Version > Valid operating system builds"
                        $SettingValue = "Not Configured"
                        $Comment = ""
                        if (!([string]::IsNullOrEmpty($CompliancePolicy.validOperatingSystemBuildRanges))) {
                            $Status = "Unsupported"
                            $SettingValue = "Found " + ($CompliancePolicy.validOperatingSystemBuildRanges).count + " valid OS configured build(s)"
                            $Comment = $URLSupportedCompliancePolicies
                            $PolicyErrors++
                        }
                        else {
                            $Status = "Supported"
                        }
                        $SettingPSObj =  [PSCustomObject]@{
                            PolicyName              = $CompliancePolicy.displayName
                            PolicyType              = $CPType
                            Setting                 = $Setting
                            Value                   = $SettingValue
                            TeamsDevicesStatus      = $Status 
                            Comment                 = $Comment
                            SettingDescription      = $SettingDescription
                            AssignedToGroup         = $outAssignedToGroup
                            ExcludedFromGroup       = $outExcludedFromGroup 
                            AssignedToGroupList     = $AssignedToGroup
                            ExcludedFromGroupList   = $ExcludedFromGroup
                            PolicyID                = $CompliancePolicy.id
                            ID                      = $ID
                        }
                        $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicyDetailed')
                        [void]$output.Add($SettingPSObj)
                        #endregion

                        #region 20: System Security > Defender > Microsoft Defender Antimalware minimum version
                        $ID = 20
                        $Setting = "defenderVersion"
                        $SettingDescription = "System Security > Defender > Microsoft Defender Antimalware minimum version"
                        $SettingValue = "Not Configured"
                        $Comment = ""
                        if (!([string]::IsNullOrEmpty($CompliancePolicy.defenderVersion))) {
                            $Status = "Unsupported"
                            $SettingValue = $CompliancePolicy.defenderVersion
                            $Comment = "Teams Rooms automatically updates this component so there's no need to set compliance policies."
                            $PolicyErrors++
                        }
                        else {
                            $Status = "Supported"
                        }
                        $SettingPSObj =  [PSCustomObject]@{
                            PolicyName              = $CompliancePolicy.displayName
                            PolicyType              = $CPType
                            Setting                 = $Setting
                            Value                   = $SettingValue
                            TeamsDevicesStatus      = $Status 
                            Comment                 = $Comment
                            SettingDescription      = $SettingDescription
                            AssignedToGroup         = $outAssignedToGroup
                            ExcludedFromGroup       = $outExcludedFromGroup 
                            AssignedToGroupList     = $AssignedToGroup
                            ExcludedFromGroupList   = $ExcludedFromGroup
                            PolicyID                = $CompliancePolicy.id
                            ID                      = $ID
                        }
                        $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicyDetailed')
                        [void]$output.Add($SettingPSObj)
                        #endregion
                    }

                    if ($PolicyErrors -gt 0) {
                        $StatusSum = "Found " + $PolicyErrors + " unsupported settings."
                        $displayWarning = $true
                    }
                    elseif ($PolicyWarnings -gt 0) {
                        $StatusSum = "Found " + $PolicyWarnings + " settings that may impact users."
                        $displayWarning = $true
                    }
                    else {
                        $StatusSum = "No issues found."
                    }
                    $PolicySum =  [PSCustomObject]@{
                        PolicyID                = $CompliancePolicy.id
                        PolicyName              = $CompliancePolicy.displayName
                        PolicyType              = $CPType
                        AssignedToGroup         = $outAssignedToGroup
                        AssignedToGroupList     = $AssignedToGroup
                        ExcludedFromGroup       = $outExcludedFromGroup 
                        ExcludedFromGroupList   = $ExcludedFromGroup
                        TeamsDevicesStatus      = $StatusSum
                    }
                    $PolicySum.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicy')
                    $outputSum.Add($PolicySum) | Out-Null
                } elseif (($AssignedToGroup.count -eq 0) -and !($UserUPN -or $DeviceID -or $Detailed)) {
                    $skippedCompliancePolicies++
                }
            }
        }
        if($totalCompliancePolicies -eq 0){
            if($UserUPN){
                Write-Warning ("The user " + $UserUPN + " doesn't have any Compliance Policies assigned.")
            } else {
                Write-Warning "No Compliance Policies assigned to All Users, All Devices or group found. Please use Test-UcTeamsDevicesCompliancePolicy -All to check all policies."
            }
        }
        if($IncludeSupported -and $Detailed)
        {
            $output | Sort-Object PolicyName, ID
        } elseif ($Detailed) {
            if ((( $output | Where-Object -Property TeamsDevicesStatus -NE -Value "Supported").count -eq 0) -and !$IncludeSupported){
                Write-Warning "No unsupported settings found, please use Test-UcTeamsDevicesCompliancePolicy -IncludeSupported to output all settings."
            } else {
                $output | Where-Object -Property TeamsDevicesStatus -NE -Value "Supported" | Sort-Object PolicyName, ID
            }
        }
        else {
            if(($skippedCompliancePolicies -gt 0) -and !$All){
                Write-Warning ("Skipping $skippedCompliancePolicies compliance policies since will not be applied to Teams Devices.")
                Write-Warning ("Please use the All switch to check all policies: Test-UcTeamsDevicesCompliancePolicy -All")
            }
            if($displayWarning){
                Write-Warning "One or more policies contain unsupported settings, please use Test-UcTeamsDevicesCompliancePolicy -Detailed to identify the unsupported settings."
            }
            $outputSum | Sort-Object PolicyName
        }
    }
}