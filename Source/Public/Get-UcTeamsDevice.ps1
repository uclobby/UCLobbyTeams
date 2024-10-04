function Get-UcTeamsDevice {
    param(
        [ValidateSet("Phone", "MTR", "MTRA", "MTRW", "SurfaceHub", "Display", "Panel", "SIPPhone")]
        [string]$Filter,
        [switch]$Detailed,
        [switch]$ExportCSV,
        [string]$OutputPath
    )    
    <#
        .SYNOPSIS
        Get Microsoft Teams Devices information

        .DESCRIPTION
        This function fetch Teams Devices provisioned in a M365 Tenant using MS Graph.
        

        Contributors: David Paulino, Silvio Schanz, Gonçalo Sepulveda, Bryan Kendrick and Daniel Jelinek

        Requirements:   Microsoft Graph PowerShell Module (Install-Module Microsoft.Graph)
                        Microsoft Graph Scopes:
                                "TeamworkDevice.Read.All"
                                "User.Read.All"

        .PARAMETER Filter
        Specifies a filter, valid options:
            Phone - Teams Native Phones
            MTR - Microsoft Teams Rooms running Windows or Android
            MTRW - Microsoft Teams Room Running Windows
            MTRA - Microsoft Teams Room Running Android
            SurfaceHub - Surface Hub 
            Display - Microsoft Teams Displays 
            Panel - Microsoft Teams Panels

        .PARAMETER Detailed
        When present it will get detailed information from Teams Devices

        .PARAMETER ExportCSV
        When present will export the detailed results to a CSV file. By defautl will save the file under the current user downloads, unless we specify the OutputPath.

        .PARAMETER OutputPath
        Allows to specify the path where we want to save the results.

        .EXAMPLE 
        PS> Get-UcTeamsDevice

        .EXAMPLE 
        PS> Get-UcTeamsDevice -Filter MTR

        .EXAMPLE
        PS> Get-UcTeamsDevice -Detailed

    #>

    $outTeamsDevices = [System.Collections.ArrayList]::new()

    if ($ExportCSV) {
        $Detailed = $true
    }

    #Verify if the Output Path exists
    if ($OutputPath) {
        if (!(Test-Path $OutputPath -PathType Container)) {
            Write-Host ("Error: Invalid folder " + $OutputPath) -ForegroundColor Red
            return
        } 
    }
    else {                
        $OutputPath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads")
    }

    if (Test-UcMgGraphConnection -Scopes "TeamworkDevice.Read.All,User.Read.All" -AltScopes ("TeamworkDevice.Read.All","Directory.Read.All")) {
        Test-UcPowerShellModule -ModuleName UcLobbyTeams | Out-Null
        $graphRequests = [System.Collections.ArrayList]::new()
        $tmpFileName = "MSTeamsDevices_" + $Filter + "_" + ( get-date ).ToString('yyyyMMdd-HHmmss') + ".csv"
        switch ($filter) {
            "Phone" { 
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "ipPhone"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'ipPhone'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "lowCostPhone"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'lowCostPhone'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
            }
            "MTR" {
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "teamsRoom"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'teamsRoom'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "collaborationBar"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'collaborationBar'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "touchConsole"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'touchConsole'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
            }
            "MTRW" {
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "teamsRoom"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'teamsRoom'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
            }
            "MTRA" {            
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "collaborationBar"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'collaborationBar'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "touchConsole"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'touchConsole'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
            }
            "SurfaceHub" {
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "surfaceHub"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'surfaceHub'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
            }
            "Display" {
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "teamsDisplay"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'teamsDisplay'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
            }
            "Panel" {
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "teamsPanel"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'teamsPanel'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
            }
            "SIPPhone" {
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "sip"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'sip'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
            }
            Default {
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = 1
                    method = "GET"
                    url    = "/teamwork/devices"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
                $tmpFileName = "MSTeamsDevices_All_" + ( get-date ).ToString('yyyyMMdd-HHmmss') + ".csv"
            }
        }
        
        $TeamsDeviceList = (Invoke-UcMgGraphBatch -Requests $graphRequests -MgProfile beta -Activity "Get-UcTeamsDevice, getting Teams device info").value
        
        #To improve performance we will use batch requests
        $graphRequests = [System.Collections.ArrayList]::new()
        foreach ($TeamsDevice in $TeamsDeviceList) {
            if (($graphRequests.id -notcontains $TeamsDevice.currentuser.id) -and !([string]::IsNullOrEmpty($TeamsDevice.currentuser.id))) {
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = $TeamsDevice.currentuser.id
                    method = "GET"
                    url    = "/users/" + $TeamsDevice.currentuser.id
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
            }
            if ($Detailed) {
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = $TeamsDevice.id + "-activity"
                    method = "GET"
                    url    = "/teamwork/devices/" + $TeamsDevice.id + "/activity"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = $TeamsDevice.id + "-configuration"
                    method = "GET"
                    url    = "/teamwork/devices/" + $TeamsDevice.id + "/configuration"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = $TeamsDevice.id + "-health"
                    method = "GET"
                    url    = "/teamwork/devices/" + $TeamsDevice.id + "/health"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = $TeamsDevice.id + "-operations"
                    method = "GET"
                    url    = "/teamwork/devices/" + $TeamsDevice.id + "/operations"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
            } 
        }
        if ($graphRequests.Count -gt 0) {
            
            if ($Detailed) {
                $ActivityInfo = "Get-UcTeamsDevice, getting Teams device addtional information (User UPN/Health/Operations/Configurarion)."
            }
            else {
                $ActivityInfo = "Get-UcTeamsDevice, getting Teams device user information."
            }
            $graphResponseExtra = (Invoke-UcMgGraphBatch -Requests $graphRequests -MgProfile beta -Activity $ActivityInfo -IncludeBody)
        }
        $devicesProcessed = 0
        foreach ($TeamsDevice in $TeamsDeviceList) {
            $devicesProcessed++
            $userUPN = ($graphResponseExtra | Where-Object { $_.id -eq $TeamsDevice.currentuser.id }).body.userPrincipalName

            if ($Detailed) {
                $TeamsDeviceActivity = ($graphResponseExtra | Where-Object { $_.id -eq ($TeamsDevice.id + "-activity") }).body
                $TeamsDeviceConfiguration = ($graphResponseExtra | Where-Object { $_.id -eq ($TeamsDevice.id + "-configuration") }).body
                $TeamsDeviceHealth = ($graphResponseExtra | Where-Object { $_.id -eq ($TeamsDevice.id + "-health") }).body
                $TeamsDeviceOperations = ($graphResponseExtra | Where-Object { $_.id -eq ($TeamsDevice.id + "-operations") }).body.value

                if ($TeamsDeviceOperations.count -gt 0) {
                    $LastHistoryAction = $TeamsDeviceOperations[0].operationType
                    $LastHistoryStatus = $TeamsDeviceOperations[0].status
                    $LastHistoryInitiatedBy = $TeamsDeviceOperations[0].createdBy.user.displayName
                    $LastHistoryModifiedDate = ($TeamsDeviceOperations[0].lastActionDateTime).ToLocalTime()
                    $LastHistoryErrorCode = $TeamsDeviceOperations[0].error.code
                    $LastHistoryErrorMessage = $TeamsDeviceOperations[0].error.message
                }
                else {
                    $LastHistoryAction = ""
                    $LastHistoryStatus = ""
                    $LastHistoryInitiatedBy = ""
                    $LastHistoryModifiedDate = ""
                    $LastHistoryErrorCode = ""
                    $LastHistoryErrorMessage = ""
                }

                $outMacAddress = ""
                foreach ($macAddress in $TeamsDevice.hardwaredetail.macAddresses) {
                    $outMacAddress += $macAddress + ";"
                }
        
                $TDObj = New-Object -TypeName PSObject -Property @{
                    UserDisplayName               = $TeamsDevice.currentuser.displayName
                    UserUPN                       = $userUPN 
            
                    TACDeviceID                   = $TeamsDevice.id
                    DeviceType                    = Convert-UcTeamsDeviceType $TeamsDevice.deviceType
                    Notes                         = $TeamsDevice.notes
                    CompanyAssetTag               = $TeamsDevice.companyAssetTag
        
                    Manufacturer                  = $TeamsDevice.hardwaredetail.manufacturer
                    Model                         = $TeamsDevice.hardwaredetail.model
                    SerialNumber                  = $TeamsDevice.hardwaredetail.serialNumber 
                    MacAddresses                  = $outMacAddress
                                
                    DeviceHealth                  = $TeamsDevice.healthStatus
                    WhenCreated                   = ($TeamsDevice.createdDateTime).ToLocalTime()
                    WhenChanged                   = ($TeamsDevice.lastModifiedDateTime).ToLocalTime()
                    ChangedByUser                 = $TeamsDevice.lastModifiedBy.user.displayName
        
                    #Activity
                    ActivePeripherals             = $TeamsDeviceActivity.activePeripherals
        
                    #Configuration
                    ConfigurationCreateDate       = ($TeamsDeviceConfiguration.createdDateTime).ToLocalTime()
                    ConfigurationCreatedBy        = $TeamsDeviceConfiguration.createdBy
                    ConfigurationLastModifiedDate = ($TeamsDeviceConfiguration.lastModifiedDateTime).ToLocalTime()
                    ConfigurationLastModifiedBy   = $TeamsDeviceConfiguration.lastModifiedBy
                    DisplayConfiguration          = $TeamsDeviceConfiguration.displayConfiguration
                    CameraConfiguration           = $TeamsDeviceConfiguration.cameraConfiguration.contentCameraConfiguration
                    SpeakerConfiguration          = $TeamsDeviceConfiguration.speakerConfiguration
                    MicrophoneConfiguration       = $TeamsDeviceConfiguration.microphoneConfiguration
                    TeamsClientConfiguration      = $TeamsDeviceConfiguration.teamsClientConfiguration
                    SupportedMeetingMode          = $TeamsDeviceConfiguration.teamsClientConfiguration.accountConfiguration.supportedClient
                    HardwareProcessor             = $TeamsDeviceConfiguration.hardwareConfiguration.processorModel
                    SystemConfiguration           = $TeamsDeviceConfiguration.systemConfiguration
        
                    #Health
                    #20240417 - Added connection fields
                    ConnectionStatus              = $TeamsDeviceHealth.connection.connectionStatus
                    ConnectionLastActivity        = ($TeamsDeviceHealth.connection.lastModifiedDateTime).ToLocalTime()
                    
                    ComputeStatus                 = $TeamsDeviceHealth.hardwareHealth.computeHealth.connection.connectionStatus
                    HdmiIngestStatus              = $TeamsDeviceHealth.hardwareHealth.hdmiIngestHealth.connection.connectionStatus
                    RoomCameraStatus              = $TeamsDeviceHealth.peripheralsHealth.roomCameraHealth.connection.connectionStatus
                    ContentCameraStatus           = $TeamsDeviceHealth.peripheralsHealth.contentCameraHealth.connection.connectionStatus
                    SpeakerStatus                 = $TeamsDeviceHealth.peripheralsHealth.speakerHealth.connection.connectionStatus
                    CommunicationSpeakerStatus    = $TeamsDeviceHealth.peripheralsHealth.communicationSpeakerHealth.connection.connectionStatus
                    #DisplayCollection = $TeamsDeviceHealth.peripheralsHealth.displayHealthCollection.connectionStatus
                    MicrophoneStatus              = $TeamsDeviceHealth.peripheralsHealth.microphoneHealth.connection.connectionStatus

                    TeamsAdminAgentVersion        = $TeamsDeviceHealth.softwareUpdateHealth.adminAgentSoftwareUpdateStatus.currentVersion
                    FirmwareVersion               = $TeamsDeviceHealth.softwareUpdateHealth.firmwareSoftwareUpdateStatus.currentVersion
                    CompanyPortalVersion          = $TeamsDeviceHealth.softwareUpdateHealth.companyPortalSoftwareUpdateStatus.currentVersion
                    OEMAgentAppVersion            = $TeamsDeviceHealth.softwareUpdateHealth.partnerAgentSoftwareUpdateStatus.currentVersion
                    TeamsAppVersion               = $TeamsDeviceHealth.softwareUpdateHealth.teamsClientSoftwareUpdateStatus.currentVersion
                    
                    #LastOperation
                    LastHistoryAction             = $LastHistoryAction
                    LastHistoryStatus             = $LastHistoryStatus
                    LastHistoryInitiatedBy        = $LastHistoryInitiatedBy
                    LastHistoryModifiedDate       = $LastHistoryModifiedDate
                    LastHistoryErrorCode          = $LastHistoryErrorCode
                    LastHistoryErrorMessage       = $LastHistoryErrorMessage 
                }
                $TDObj.PSObject.TypeNames.Insert(0, 'TeamsDevice')
        
            }
            else {
                $TDObj = New-Object -TypeName PSObject -Property @{
                    UserDisplayName = $TeamsDevice.currentuser.displayName
                    UserUPN         = $userUPN 
                    TACDeviceID     = $TeamsDevice.id
                    DeviceType      = Convert-UcTeamsDeviceType $TeamsDevice.deviceType
                    Manufacturer    = $TeamsDevice.hardwaredetail.manufacturer
                    Model           = $TeamsDevice.hardwaredetail.model
                    SerialNumber    = $TeamsDevice.hardwaredetail.serialNumber 
                    MacAddresses    = $TeamsDevice.hardwaredetail.macAddresses
                                
                    DeviceHealth    = $TeamsDevice.healthStatus

                    #20240419 - Adding additional fields that are available on graph api
                    WhenCreated     = ($TeamsDevice.createdDateTime).ToLocalTime()
                    WhenChanged     = ($TeamsDevice.lastModifiedDateTime).ToLocalTime()
                    ChangedByUser   = $TeamsDevice.lastModifiedBy.user.displayName
                }
                $TDObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceList')
            }
            $outTeamsDevices.Add($TDObj) | Out-Null
        }
        #20231020 - We only need to output if we have results.
        if ($devicesProcessed -gt 0) {
            #region: Modified by Daniel Jelinek
            if ($ExportCSV) {
                $OutputFullPath = [System.IO.Path]::Combine($OutputPath, $tmpFileName)
                $outTeamsDevices | Sort-Object DeviceType, Manufacturer, Model | Select-Object TACDeviceID, DeviceType, Manufacturer, Model, UserDisplayName, UserUPN, Notes, CompanyAssetTag, SerialNumber, MacAddresses, WhenCreated, WhenChanged, ChangedByUser, HdmiIngestStatus, ComputeStatus, RoomCameraStatus, SpeakerStatus, CommunicationSpeakerStatus, MicrophoneStatus, SupportedMeetingMode, HardwareProcessor, SystemConfiguratio, TeamsAdminAgentVersion, FirmwareVersion, CompanyPortalVersion, OEMAgentAppVersion, TeamsAppVersion, LastUpdate, LastHistoryAction, LastHistoryStatus, LastHistoryInitiatedBy, LastHistoryModifiedDate, LastHistoryErrorCode, LastHistoryErrorMessage | Export-Csv -path $OutputFullPath -NoTypeInformation
                Write-Host ("Results available in: " + $OutputFullPath) -ForegroundColor Cyan
            }
            else {
                $outTeamsDevices | Sort-Object DeviceType, Manufacturer, Model
            }
            #endregion
        }
    }
}