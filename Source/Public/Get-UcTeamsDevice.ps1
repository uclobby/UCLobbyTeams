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

$GraphURI_BetaAPIBatch = "https://graph.microsoft.com/beta/`$batch"

Function Get-UcTeamsDevice {
    Param(
        [ValidateSet("Phone","MTR","MTRA","MTRW","SurfaceHub","Display","Panel","SIPPhone")]
        [string]$Filter,
        [switch]$Detailed,
        [switch]$ExportCSV,
        [string]$OutputPath
    )    
    $outTeamsDevices = [System.Collections.ArrayList]::new()
    $TeamsDeviceList =  [System.Collections.ArrayList]::new()
    $graphRequestsExtra =  [System.Collections.ArrayList]::new()
    $graphResponseExtra =  [System.Collections.ArrayList]::new()

    if($ExportCSV){
        $Detailed = $true
    }

    #Verify if the Output Path exists
    if($OutputPath){
        if (!(Test-Path $OutputPath -PathType Container)){
            Write-Host ("Error: Invalid folder " + $OutputPath) -ForegroundColor Red
            return
        } 
    } else {                
        $OutputPath = [System.IO.Path]::Combine($env:USERPROFILE,"Downloads")
    }

    if(Test-UcMgGraphConnection -Scopes "TeamworkDevice.Read.All", "User.Read.All"){
        $graphRequests =  [System.Collections.ArrayList]::new()
        $tmpFileName = "MSTeamsDevices_" + $Filter + "_" + ( get-date ).ToString('yyyyMMdd-HHmmss') + ".csv"
        switch ($filter) {
            "Phone" { 
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id = "ipPhone"
                    method = "GET"
                    url = "/teamwork/devices/?`$filter=deviceType eq 'ipPhone'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id = "lowCostPhone"
                    method = "GET"
                    url = "/teamwork/devices/?`$filter=deviceType eq 'lowCostPhone'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
            }
            "MTR" {
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id = "teamsRoom"
                    method = "GET"
                    url = "/teamwork/devices/?`$filter=deviceType eq 'teamsRoom'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id = "collaborationBar"
                    method = "GET"
                    url = "/teamwork/devices/?`$filter=deviceType eq 'collaborationBar'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id = "touchConsole"
                    method = "GET"
                    url = "/teamwork/devices/?`$filter=deviceType eq 'touchConsole'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
            }
            "MTRW"{
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id = "teamsRoom"
                    method = "GET"
                    url = "/teamwork/devices/?`$filter=deviceType eq 'teamsRoom'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
            }
            "MTRA"{            
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id = "collaborationBar"
                    method = "GET"
                    url = "/teamwork/devices/?`$filter=deviceType eq 'collaborationBar'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id = "touchConsole"
                    method = "GET"
                    url = "/teamwork/devices/?`$filter=deviceType eq 'touchConsole'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
            }
            "SurfaceHub" {
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id = "surfaceHub"
                    method = "GET"
                    url = "/teamwork/devices/?`$filter=deviceType eq 'surfaceHub'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
            }
            "Display"{
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id = "teamsDisplay"
                    method = "GET"
                    url = "/teamwork/devices/?`$filter=deviceType eq 'teamsDisplay'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
            }
            "Panel" {
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id = "teamsPanel"
                    method = "GET"
                    url = "/teamwork/devices/?`$filter=deviceType eq 'teamsPanel'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
            }
            "SIPPhone" {
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id = "sip"
                    method = "GET"
                    url = "/teamwork/devices/?`$filter=deviceType eq 'sip'"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
            }
            Default {
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id = 1
                    method = "GET"
                    url = "/teamwork/devices"
                }
                $graphRequests.Add($gRequestTmp) | Out-Null
                $tmpFileName = "MSTeamsDevices_All_" + ( get-date ).ToString('yyyyMMdd-HHmmss') + ".csv"
            }
        }
        
        #TO DO: Look for alternatives instead of doing this.
        if($graphRequests.Count -gt 1){
            $graphBody = ' { "requests":  '+ ($graphRequests | ConvertTo-Json) + ' }' 
        } else {
            $graphBody = ' { "requests": ['+ ($graphRequests | ConvertTo-Json) + '] }' 
        }
        
        $tmpGraphResponses = (Invoke-MgGraphRequest -Method Post -Uri $GraphURI_BetaAPIBatch -Body $graphBody).responses
        for($j=0;$j -lt $tmpGraphResponses.length; $j++){
            if($tmpGraphResponses[$j].status -eq 200){  
                $TeamsDeviceList += $tmpGraphResponses[$j].body.value
                #Checking if there are more pages available
                $GraphURI_NextPage = $tmpGraphResponses[$j].body.'@odata.nextLink'
                while(![string]::IsNullOrEmpty($GraphURI_NextPage)){
                    $graphNextPageResponse =  Invoke-MgGraphRequest -Method Get -Uri $GraphURI_NextPage
                    $TeamsDeviceList  += $graphNextPageResponse.value
                    $GraphURI_NextPage = $graphNextPageResponse.'@odata.nextLink'
                }
            }
        }
        
        #To improve performance we will use batch requests
        $i = 1
        foreach($TeamsDevice in $TeamsDeviceList){
            $batchCount = [int](($TeamsDeviceList.length * 5)/20)+1
            Write-Progress -Activity "Teams Device List" -Status "Running batch $i of $batchCount"  -PercentComplete (($i / $batchCount) * 100)
            if(($graphRequestsExtra.id -notcontains $TeamsDevice.currentuser.id) -and !([string]::IsNullOrEmpty($TeamsDevice.currentuser.id)) -and ($graphResponseExtra.id -notcontains $TeamsDevice.currentuser.id)) {
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id =  $TeamsDevice.currentuser.id
                    method = "GET"
                    url = "/users/"+ $TeamsDevice.currentuser.id
                }
                $graphRequestsExtra.Add($gRequestTmp) | Out-Null
            }
            if($Detailed){
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id = $TeamsDevice.id+"-activity"
                    method = "GET"
                    url = "/teamwork/devices/"+$TeamsDevice.id+"/activity"
                }
                $graphRequestsExtra.Add($gRequestTmp) | Out-Null
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id = $TeamsDevice.id+"-configuration"
                    method = "GET"
                    url = "/teamwork/devices/"+$TeamsDevice.id+"/configuration"
                }
                $graphRequestsExtra.Add($gRequestTmp) | Out-Null
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id =$TeamsDevice.id+"-health"
                    method = "GET"
                    url = "/teamwork/devices/"+$TeamsDevice.id+"/health"
                }
                $graphRequestsExtra.Add($gRequestTmp) | Out-Null
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id = $TeamsDevice.id+"-operations"
                    method = "GET"
                    url = "/teamwork/devices/"+$TeamsDevice.id+"/operations"
                }
                $graphRequestsExtra.Add($gRequestTmp) | Out-Null
            } 

            #MS Graph is limited to 20 requests per batch, each device has 5 requests unless we already know the User UPN.
            if($graphRequestsExtra.Count -gt 15)  {
                $i++
                $graphBodyExtra = ' { "requests":  '+ ($graphRequestsExtra  | ConvertTo-Json) + ' }' 
                $graphResponseExtra += (Invoke-MgGraphRequest -Method Post -Uri $GraphURI_BetaAPIBatch -Body $graphBodyExtra).responses
                $graphRequestsExtra =  [System.Collections.ArrayList]::new()
            }
        }
        if ($graphRequestsExtra.Count -gt 0){
            Write-Progress -Activity "Teams Device List" -Status "Running batch $i of $batchCount"  -PercentComplete (($i / $batchCount) * 100)
            if($graphRequestsExtra.Count -gt 1){
                $graphBodyExtra = ' { "requests":  '+ ($graphRequestsExtra | ConvertTo-Json) + ' }' 
            } else {
                $graphBodyExtra = ' { "requests": ['+ ($graphRequestsExtra | ConvertTo-Json) + '] }' 
            }
            $graphResponseExtra += (Invoke-MgGraphRequest -Method Post -Uri $GraphURI_BetaAPIBatch -Body $graphBodyExtra).responses
        }
        
        foreach($TeamsDevice in $TeamsDeviceList){
            $userUPN = ($graphResponseExtra | Where-Object{$_.id -eq $TeamsDevice.currentuser.id}).body.userPrincipalName

            if($Detailed){
                $TeamsDeviceActivity = ($graphResponseExtra | Where-Object{$_.id -eq ($TeamsDevice.id+"-activity")}).body
                $TeamsDeviceConfiguration = ($graphResponseExtra | Where-Object{$_.id -eq ($TeamsDevice.id+"-configuration")}).body
                $TeamsDeviceHealth = ($graphResponseExtra | Where-Object{$_.id -eq ($TeamsDevice.id+"-health")}).body
                $TeamsDeviceOperations = ($graphResponseExtra | Where-Object{$_.id -eq ($TeamsDevice.id+"-operations")}).body.value

                if($TeamsDeviceOperations.count -gt 0){
                    $LastHistoryAction = $TeamsDeviceOperations[0].operationType
                    $LastHistoryStatus = $TeamsDeviceOperations[0].status
                    $LastHistoryInitiatedBy = $TeamsDeviceOperations[0].createdBy.user.displayName
                    $LastHistoryModifiedDate = $TeamsDeviceOperations[0].lastActionDateTime
                    $LastHistoryErrorCode = $TeamsDeviceOperations[0].error.code
                    $LastHistoryErrorMessage = $TeamsDeviceOperations[0].error.message
                } else {
                    $LastHistoryAction = ""
                    $LastHistoryStatus = ""
                    $LastHistoryInitiatedBy = ""
                    $LastHistoryModifiedDate = ""
                    $LastHistoryErrorCode = ""
                    $LastHistoryErrorMessage = ""
                }

                $outMacAddress = ""
                foreach ($macAddress in $TeamsDevice.hardwaredetail.macAddresses){
                    $outMacAddress += $macAddress + ";"
                }
        
                $TDObj = New-Object -TypeName PSObject -Property @{
                    UserDisplayName = $TeamsDevice.currentuser.displayName
                    UserUPN         = $userUPN 
            
                    TACDeviceID     = $TeamsDevice.id
                    DeviceType      = Convert-UcTeamsDeviceType $TeamsDevice.deviceType
                    Notes           = $TeamsDevice.notes
                    CompanyAssetTag = $TeamsDevice.companyAssetTag
        
                    Manufacturer    = $TeamsDevice.hardwaredetail.manufacturer
                    Model           = $TeamsDevice.hardwaredetail.model
                    SerialNumber    = $TeamsDevice.hardwaredetail.serialNumber 
                    MacAddresses    = $outMacAddress.subString(0,$outMacAddress.length-1)
                                
                    DeviceHealth    = $TeamsDevice.healthStatus
                    WhenCreated = $TeamsDevice.createdDateTime
                    WhenChanged = $TeamsDevice.lastModifiedDateTime
                    ChangedByUser = $TeamsDevice.lastModifiedBy.user.displayName
        
                    #Activity
                    ActivePeripherals = $TeamsDeviceActivity.activePeripherals
        
                    #Configuration
                    LastUpdate = $TeamsDeviceConfiguration.createdDateTime
        
                    DisplayConfiguration = $TeamsDeviceConfiguration.displayConfiguration
                    CameraConfiguration = $TeamsDeviceConfiguration.cameraConfiguration.contentCameraConfiguration
                    SpeakerConfiguration = $TeamsDeviceConfiguration.speakerConfiguration
                    MicrophoneConfiguration = $TeamsDeviceConfiguration.microphoneConfiguration
                    TeamsClientConfiguration = $TeamsDeviceConfiguration.teamsClientConfiguration
                    SupportedMeetingMode = $TeamsDeviceConfiguration.teamsClientConfiguration.accountConfiguration.supportedClient
                    HardwareProcessor = $TeamsDeviceConfiguration.hardwareConfiguration.processorModel
                    SystemConfiguration = $TeamsDeviceConfiguration.systemConfiguration
        
                    #Health
                    ComputeStatus = $TeamsDeviceHealth.hardwareHealth.computeHealth.connection.connectionStatus
                    HdmiIngestStatus = $TeamsDeviceHealth.hardwareHealth.hdmiIngestHealth.connection.connectionStatus
                    RoomCameraStatus = $TeamsDeviceHealth.peripheralsHealth.roomCameraHealth.connection.connectionStatus
                    ContentCameraStatus = $TeamsDeviceHealth.peripheralsHealth.contentCameraHealth.connection.connectionStatus
                    SpeakerStatus = $TeamsDeviceHealth.peripheralsHealth.speakerHealth.connection.connectionStatus
                    CommunicationSpeakerStatus = $TeamsDeviceHealth.peripheralsHealth.communicationSpeakerHealth.connection.connectionStatus
                    #DisplayCollection = $TeamsDeviceHealth.peripheralsHealth.displayHealthCollection.connectionStatus
                    MicrophoneStatus = $TeamsDeviceHealth.peripheralsHealth.microphoneHealth.connection.connectionStatus

                    TeamsAdminAgentVersion = $TeamsDeviceHealth.softwareUpdateHealth.adminAgentSoftwareUpdateStatus.currentVersion
                    FirmwareVersion = $TeamsDeviceHealth.softwareUpdateHealth.firmwareSoftwareUpdateStatus.currentVersion
                    CompanyPortalVersion = $TeamsDeviceHealth.softwareUpdateHealth.companyPortalSoftwareUpdateStatus.currentVersion
                    OEMAgentAppVersion = $TeamsDeviceHealth.softwareUpdateHealth.partnerAgentSoftwareUpdateStatus.currentVersion
                    TeamsAppVersion = $TeamsDeviceHealth.softwareUpdateHealth.teamsClientSoftwareUpdateStatus.currentVersion
                    
                    #LastOperation
                    LastHistoryAction = $LastHistoryAction
                    LastHistoryStatus = $LastHistoryStatus
                    LastHistoryInitiatedBy = $LastHistoryInitiatedBy
                    LastHistoryModifiedDate = $LastHistoryModifiedDate
                    LastHistoryErrorCode = $LastHistoryErrorCode
                    LastHistoryErrorMessage = $LastHistoryErrorMessage 
                }
                $TDObj.PSObject.TypeNames.Insert(0, 'TeamsDevice')
        
            } else {
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
                }
                $TDObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceList')
            }
            $outTeamsDevices.Add($TDObj) | Out-Null
        }
        #region: Modified by Daniel Jelinek
        if($ExportCSV){
            $OutputFullPath = [System.IO.Path]::Combine($OutputPath, $tmpFileName)
            $outTeamsDevices | Sort-Object DeviceType,Manufacturer,Model| Select-Object TACDeviceID, DeviceType, Manufacturer, Model, UserDisplayName, UserUPN, Notes, CompanyAssetTag, SerialNumber, MacAddresses, WhenCreated, WhenChanged, ChangedByUser, HdmiIngestStatus, ComputeStatus, RoomCameraStatus, SpeakerStatus, CommunicationSpeakerStatus, MicrophoneStatus, SupportedMeetingMode, HardwareProcessor, SystemConfiguratio, TeamsAdminAgentVersion, FirmwareVersion, CompanyPortalVersion, OEMAgentAppVersion, TeamsAppVersion, LastUpdate, LastHistoryAction, LastHistoryStatus, LastHistoryInitiatedBy, LastHistoryModifiedDate, LastHistoryErrorCode, LastHistoryErrorMessage| Export-Csv -path $OutputFullPath -NoTypeInformation
            Write-Host ("Results available in: " + $OutputFullPath) -ForegroundColor Cyan
        } else {
            $outTeamsDevices | Sort-Object DeviceType,Manufacturer,Model
        }
        #endregion
    }
}