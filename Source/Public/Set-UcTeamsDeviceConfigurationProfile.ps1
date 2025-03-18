function Set-UcTeamsDeviceConfigurationProfile {
    <#
        .SYNOPSIS
        Allow assign a Teams Device Configuration Profile to one or more Teams Devices

        .DESCRIPTION
        This function will use TAC API to assign a Configuration Profile by sending the Config Update command to the specified device(s).

        Requirements:   EntraAuth PowerShell Module (Install-Module EntraAuth)

        .PARAMETER TACDeviceID
        Teams Device ID from Teams Admin Center

        .PARAMETER ConfigID
        Teams Device Configuration Profile ID

        .EXAMPLE 
        PS> Set-UcTeamsDeviceConfigurationProfile -TACDeviceID "00000000-0000-0000-0000-000000000000" -ConfigID "00000000-0000-0000-0000-000000000000"
    #>
    param(
        [Parameter(mandatory=$true)]    
        [string[]]$TACDeviceID,
        [Parameter(mandatory=$true)]    
        [string]$ConfigID
    )
    if (!($global:UCLobbyTeamsModuleCheck)) {
        Test-UcPowerShellModule -ModuleName UcLobbyTeams | Out-Null
        $global:UCLobbyTeamsModuleCheck = $true
    }
    if (Test-UcAPIConnection -Type TeamsDeviceTAC) {
        $teamsDevices = [System.Collections.ArrayList]::new()
        $cmdUpdates = [System.Collections.ArrayList]::new()
        $output = [System.Collections.ArrayList]::new()
        #Checking if the Configuration Profile is valid, we also need to make sure the DeviceType match the Teams Devices.
        $configProfile = Get-UcTeamsDeviceConfigurationProfile -Identity $ConfigID
        foreach ($singleTACDeviceID in $TACDeviceID) {
            $TeamsDeviceInfo = Get-UcTeamsDevice -UseTAC -TACDeviceID $singleTACDeviceID
            if ($TeamsDeviceInfo) {              
                if ($TeamsDeviceInfo.DeviceType -eq $configProfile.DeviceType) {
                    #We need to confirm that we don't have pending Configuration Profile updates.
                    $cmdHistory = (Invoke-EntraRequest -Path "api/v2/devices/$singleTACDeviceID/commands" -Service TeamsDeviceTAC).commands
                    if ($cmdHistory | Where-Object {$_.command -eq "ConfigUpdate" -and $_.commandStatus -eq "Queued"} ){
                        Write-Warning "Skipping Device $singleTACDeviceID because already has a Configuration Profile Update queued."
                    } else {
                        $cmdUpdateObj = [Ordered]@{
                            'device-id' = $singleTACDeviceID 
                            'Commands'  = @(@{
                                    'cmd'       = 'ConfigUpdate'
                                    'payloadId' = $ConfigID
                                })
                        }
                        [void]$cmdUpdates.Add($cmdUpdateObj)
                        [void]$teamsDevices.Add($TeamsDeviceInfo)
                    }
                }
                else {
                    Write-Warning ("Skiping TACDeviceID $singleTACDeviceID, Device Type (" + $TeamsDeviceInfo.DeviceType + ") doesn't match Configuration Profile Device Type (" + $configProfile.DeviceType + ")" )
                }
            }
        }
        if ($cmdUpdates.Count -eq 1) {
            $requestBodyJson = "[" + ($cmdUpdates | ConvertTo-Json -Compress) + "]"
        }
        else {
            $requestBodyJson = $cmdUpdates | ConvertTo-Json -Compress -Depth 3
        }
        $requestHeader = New-Object 'System.Collections.Generic.Dictionary[string, string]'
        $requestHeader.Add("Content-Type", "application/json")
        $cmdResponses = (Invoke-EntraRequest -Path "/admin/api/v1/devices/commands" -Service TeamsDeviceTAC -Method POST -Header $requestHeader -Body $requestBodyJson).devices

        foreach ($cmdResponse in $cmdResponses) {
            $TeamsDeviceInfo = $teamsDevices | Where-Object -Property TACDeviceID -EQ -Value $cmdResponse.id
            $outputObj = [PSCustomObject][Ordered]@{
                DeviceTACID                  = $cmdResponse.Id
                Manufacturer                 = $TeamsDeviceInfo.Manufacturer
                Model                        = $TeamsDeviceInfo.Model
                PreviousConfigurationProfile = $TeamsDeviceInfo.ConfigurationProfile
                NewConfigurationProfile      = $configProfile.DisplayName
                DeviceStatus                 = $cmdResponse.deviceStatus
                ConfigurationUpdateStatus    = $cmdResponse.commandStatus
            }
            $outputObj.PSObject.TypeNames.Insert(0, 'SetTeamsDeviceConfigurationProfile')
            [void]$output.Add($outputObj)
        }
        return $output
    }
}