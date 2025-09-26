function Convert-UcTeamsDeviceAutoUpdateDays {
    param (
        [int]$Days
    )
    switch ($Days) {
        0 { return "Validation" }
        30 { return "General" }
        90 { return "Final" }
        default { return $Days }
    }
}

function Convert-UcTeamsDeviceSignInMode {
    param (
        [string]$SignInMode
    )
    switch ($SignInMode) {
        "commonArea" { return "Common Area" }
        "personal" { return "User" }
        "conference" { return "Conference" }
        default { return $SignInMode}
    }
}

function Convert-UcTeamsDeviceType {
    param (
        [string]$DeviceType
    )
    switch ($DeviceType) {
        "ipPhone" { return "Phone" }
        "lowCostPhone" { return "Phone" }
        "teamsRoom" { return "MTR Windows" }
        "collaborationBar" { return "MTR Android" }
        "surfaceHub" { return "Surface Hub" }
        "teamsDisplay" { return "Display" }
        "touchConsole" { return "Touch Console (MTRA)" }
        "teamsPanel" { return "Panel" }
        "sip" { return "SIP Phone" }
        Default { return $DeviceType}
    }
}

function ConvertTo-IPv4MaskString {
    <#
        .SYNOPSIS
        Converts a number of bits (0-32) to an IPv4 network mask string (e.g., "255.255.255.0").
    
        .DESCRIPTION
        Converts a number of bits (0-32) to an IPv4 network mask string (e.g., "255.255.255.0").
    
        .PARAMETER MaskBits
        Specifies the number of bits in the mask.

        Credits to: Bill Stewart - https://www.itprotoday.com/powershell/working-ipv4-addresses-powershell  
    #>
    param(
        [parameter(Mandatory = $true)]
        [ValidateRange(0, 32)]
        [Int] $MaskBits
    )
    
    $mask = ([Math]::Pow(2, $MaskBits) - 1) * [Math]::Pow(2, (32 - $MaskBits))
    $bytes = [BitConverter]::GetBytes([UInt32] $mask)
    (($bytes.Count - 1)..0 | ForEach-Object { [String] $bytes[$_] }) -join "."
}

function Get-UcM365Domains {
    <#
        .SYNOPSIS
        Get Microsoft 365 Domains from a Tenant

        .DESCRIPTION
        This function returns a list of domains that are associated with a Microsoft 365 Tenant.

        .PARAMETER Domain
        Specifies a domain registered with Microsoft 365.

        .EXAMPLE
        PS> Get-UcM365Domains -Domain uclobby.com
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Domain
    )

    $regex = "^(.*@)(.*[.].*)$"
    $outDomains = [System.Collections.ArrayList]::new()
    try {
        #2025-01-31: Only need to check this once per PowerShell session
        if (!($global:UCLobbyTeamsModuleCheck)) {
            Test-UcPowerShellModule -ModuleName UcLobbyTeams | Out-Null
            $global:UCLobbyTeamsModuleCheck = $true
        }
        $AllowedAudiences = Invoke-WebRequest -Uri ("https://accounts.accesscontrol.windows.net/" + $Domain + "/metadata/json/1") -UseBasicParsing | ConvertFrom-Json | Select-Object -ExpandProperty allowedAudiences
    }
    catch [System.Net.WebException] {
        if ($PSItem.Exception.Message -eq "The remote server returned an error: (400) Bad Request.") {
            Write-Warning "The domain $Domain is not part of a Microsoft 365 Tenant."
        }
        else {
            Write-Warning $PSItem.Exception.Message
        }
    }
    catch {
        #2024-03-18: Support for GCC High tenants.
        try {
            $AllowedAudiences = Invoke-WebRequest -Uri ("https://login.microsoftonline.us/" + $Domain + "/metadata/json/1") -UseBasicParsing | ConvertFrom-Json | Select-Object -ExpandProperty allowedAudiences
        }
        catch {
            Write-Warning "Unknown error while checking domain: $Domain"
        }
    }
    try {
        foreach ($AllowedAudience in $AllowedAudiences) {
            $temp = [regex]::Match($AllowedAudience , $regex).captures.groups
            if ($temp.count -ge 2) {
                $tempObj = New-Object -TypeName PSObject -Property @{
                    Name = $temp[2].value
                }
                $outDomains.Add($tempObj) | Out-Null
            }
        }
    }
    catch {
        Write-Warning "Unknown error while checking domain: $Domain"
    }
    return $outDomains
}

function Invoke-UcGraphRequest {
    <#
        .SYNOPSIS
        Invoke a Microsoft Graph Request using Entra Auth or Microsoft.Graph.Authentication 

        .DESCRIPTION
        This function will send a Microsoft Graph request to an available connections, "Test-UcServiceConnection -Type MsGraph" will have to be executed first to determine if we have a session with EntraAuth or Microsoft.Graph.Authentication.

        Requirements:   EntraAuth PowerShell module (Install-Module EntraAuth)
                        or
                        Microsoft Graph Authentication PowerShell Module (Install-Module Microsoft.Graph.Authentication)

        .PARAMETER Path
        Specifies Microsoft Graph Path that we want to send the request.

        .PARAMETER Header
        Specify the header for cases we need to have a custom header.

        .PARAMETER Requests
        If wwe want to send a batch request.

        .PARAMETER Beta
        When present, it will use the Microsoft Graph Beta API.

        .PARAMETER IncludeBody
        Some Ms Graph APIs can require specific AuthType, Application or Delegated (User).

        .PARAMETER Activity
        For Batch requests we have use this for Activity Progress.
    #>
    param(
        [string]$Path = "/`$batch",
        [object]$Header,
        [object]$Requests,
        [switch]$Beta,
        [switch]$IncludeBody,
        [string]$Activity
    )
    #This is an easy way to switch between v1.0 and beta.
    $BatchPath = "`$batch"
    if ($Beta) {
        $Path = "../beta" + $Path
        $BatchPath = "../beta/`$batch"
    }

    #If requests then we need to do a batch request to Graph.
    if (!$Requests) {
        if ($script:GraphEntraAuth) {
            if ($Header) {
                return Invoke-EntraRequest -Path $Path -NoPaging -Header $Header
            } 
            return Invoke-EntraRequest -Path $Path -NoPaging
        }
        else {
            if ($Header) {
                $GraphResponse = Invoke-MgGraphRequest -Uri ("/v1.0/" + $Path) -Headers $Header
            }
            else {
                $GraphResponse = Invoke-MgGraphRequest -Uri ("/v1.0/" + $Path)
            }
            #When it's more than one result Invoke-MgGraphRequest returns "value", we need to remove it to match EntraAuth behaviour.
            if ($GraphResponse.value) {
                return $GraphResponse.value
            }
            else {
                return $GraphResponse
            }
        }
    }
    else {
        $outBatchResponses = [System.Collections.ArrayList]::new()
        $tmpGraphRequests = [System.Collections.ArrayList]::new()
        $g = 1
        $requestHeader = New-Object 'System.Collections.Generic.Dictionary[string, string]'
        $requestHeader.Add("Content-Type", "application/json")
        #If activity is null then we can use this to get the function that call this function.
        if (!($Activity)) {
            $Activity = [string]$(Get-PSCallStack)[1].FunctionName
        }
        $batchCount = [int][Math]::Ceiling(($Requests.count / 20))
        foreach ($GraphRequest in $Requests) {
            Write-Progress -Activity $Activity -Status "Running batch $g of $batchCount"
            [void]$tmpGraphRequests.Add($GraphRequest) 
            if ($tmpGraphRequests.Count -ge 20) {
                $g++
                $grapRequestBody = ' { "requests":  ' + ($tmpGraphRequests  | ConvertTo-Json) + ' }' 
                if ($script:GraphEntraAuth) {
                    #TODO: Add support for Graph Batch with EntraAuth
                    $GraphResponses += (Invoke-EntraRequest -Path $BatchPath -Body $grapRequestBody -Method Post -Header $requestHeader).responses
                }
                else {
                    $GraphResponses += (Invoke-MgGraphRequest -Method Post -Uri ("/v1.0/" + $BatchPath) -Body $grapRequestBody).responses
                }
                $tmpGraphRequests = [System.Collections.ArrayList]::new()
            }
        }
        
        if ($tmpGraphRequests.Count -gt 0) {
            Write-Progress -Activity $Activity -Status "Running batch $g of $batchCount"
            #TO DO: Look for alternatives instead of doing this.
            if ($tmpGraphRequests.Count -gt 1) {
                $grapRequestBody = ' { "requests":  ' + ($tmpGraphRequests | ConvertTo-Json) + ' }' 
            }
            else {
                $grapRequestBody = ' { "requests": [' + ($tmpGraphRequests | ConvertTo-Json) + '] }' 
            }
            try {
                if ($script:GraphEntraAuth) {
                    #TODO: Add support for Graph Batch with EntraAuth
                    $GraphResponses += (Invoke-EntraRequest -Path $BatchPath -Body $grapRequestBody -Method Post -Header $requestHeader).responses
                }
                else {
                    $GraphResponses += (Invoke-MgGraphRequest -Method Post -Uri  ("/v1.0/" + $BatchPath) -Body $grapRequestBody).responses
                }
            }
            catch {
                Write-Warning "Error while getting the Graph Request."
            }
        }
        
        #In some cases we will need the complete graph response, in that case the calling function will have to process pending pages.
        $attempts = 1
        for ($j = 0; $j -lt $GraphResponses.length; $j++) {
            $ResponseCount = 0
            if ($IncludeBody) {
                $outBatchResponses += $GraphResponses[$j]
            }
            else {
                $outBatchResponses += $GraphResponses[$j].body
                if ($GraphResponses[$j].status -eq "200") {
                    #Checking if there are more pages available    
                    $GraphURI_NextPage = $GraphResponses[$j].body.'@odata.nextLink'
                    $GraphTotalCount = $GraphResponses[$j].body.'@odata.count'
                    $ResponseCount += $GraphResponses[$j].body.value.count
                    while (![string]::IsNullOrEmpty($GraphURI_NextPage)) {
                        try {
                            if ($script:GraphEntraAuth) {
                                #TODO: Add support for Graph Batch with EntraAuth, for now we need to use NoPaging to have the same behaviour as Invoke-MgGraphRequest
                                $graphNextPageResponse = Invoke-EntraRequest -Path $GraphURI_NextPage -NoPaging
                            }
                            else {
                                $graphNextPageResponse = Invoke-MgGraphRequest -Method Get -Uri $GraphURI_NextPage
                            }
                            $outBatchResponses += $graphNextPageResponse
                            $GraphURI_NextPage = $graphNextPageResponse.'@odata.nextLink'
                            $ResponseCount += $graphNextPageResponse.value.count
                            Write-Progress -Activity $Activity -Status "$ResponseCount of $GraphTotalCount"
                        }
                        catch {
                            Write-Warning "Failed to get the next batch page, retrying..."
                            $attempts--
                        }
                        if ($attempts -eq 0) {
                            Write-Warning "Could not get next batch page, skiping it."
                            break
                        }
                    }
                }
                else {
                    Write-Warning ("Failed to get Graph Response" + [Environment]::NewLine + `
                            "Error Code: " + $GraphResponses[$j].status + " " + $GraphResponses[$j].body.error.code + [Environment]::NewLine + `
                            "Error Message: " + $GraphResponses[$j].body.error.message + [Environment]::NewLine + `
                            "Request Date: " + $GraphResponses[$j].body.error.innerError.date + [Environment]::NewLine + `
                            "Request ID: " + $GraphResponses[$j].body.error.innerError.'request-id' + [Environment]::NewLine + `
                            "Client Request Id: " + $GraphResponses[$j].body.error.innerError.'client-request-id')
                } 
            }
        }
        return $outBatchResponses
    }
}

function Test-UcElevatedPrivileges {
    if (!(([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))) {
        return $false
    }
    return $true
}

function Test-UcIPaddressInSubnet {
    <#
        .SYNOPSIS
        Check if an IP address is part of an Subnet.

        .DESCRIPTION
        Returns true if the given IP address is part of the subnet, false for not or invalid ip address.
    
        Contributors: David Paulino

        .PARAMETER IPAddress
        IP Address that we want to confirm that belongs to a range.

        .PARAMETER Subnet
        Subnet in the IPaddress/SubnetMaskBits.

        .EXAMPLE 
        PS> Test-UcIPaddressInSubnet -IPAddress 192.168.0.1 -Subnet 192.168.0.0/24

    #>
    param(
        [Parameter(mandatory = $true)]    
        [string]$IPAddress,
        [Parameter(mandatory = $true)]
        [string]$Subnet
    )
    $regExIPAddressSubnet = "^((25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9])\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[0-9]))\/(3[0-2]|[1-2]{1}[0-9]{1}|[1-9])$"
    try {
        [void]($Subnet -match $regExIPAddressSubnet)
        $IPSubnet = [ipaddress]$Matches[1]
        $tmpIPAddress = [ipaddress]$IPAddress
        $subnetMask = ConvertTo-IPv4MaskString $Matches[6]
        $tmpSubnet = [ipaddress] ($subnetMask)
        $netidSubnet = [ipaddress]($IPSubnet.address -band $tmpSubnet.address)
        $netidIPAddress = [ipaddress]($tmpIPAddress.address -band $tmpSubnet.address)
        return ($netidSubnet.ipaddresstostring -eq $netidIPAddress.ipaddresstostring)
    }
    catch {
        return $false
    }
}

function Test-UcServiceConnection {
    <#
        .SYNOPSIS
        Test connection to a Service

        .DESCRIPTION
        This function will validate if the there is an active connection to a service and also if the required module is installed.

        Requirements:   MsGraph, TeamsDeviceTAC - EntraAuth PowerShell module (Install-Module EntraAuth)
                        TeamsModule - MicrosoftTeams PowerShell module (Install-Module MicrosoftTeams)

        .PARAMETER Type
        Specifies a Type of Service, valid options:
            MSGraph - Microsoft Graph
            TeamsModule - Microsoft Teams PowerShell module
            TeamsDeviceTAC - Teams Admin Center (TAC) API for Teams Devices

        .PARAMETER Scopes
        When present it will check if the require permissions are in the current Scope, only applicable to Microsoft Graph API.
        
        .PARAMETER AltScopes
        Allows checking for alternative permissions to the ones specified in AltScopes, only applicable to Microsoft Graph API.

        .PARAMETER AuthType
        Some Ms Graph APIs can require specific AuthType, Application or Delegated (User)
    #>
    param(
        [Parameter(mandatory = $true)]
        [ValidateSet("MSGraph", "TeamsPowerShell", "TeamsDeviceTAC")]
        [string]$Type,
        [string[]]$Scopes,
        [string[]]$AltScopes,
        [ValidateSet("Application", "Delegated")]
        [string]$AuthType
    )
    switch ($Type) {
        "MSGraph" {
            #UCLobbyTeams is moving to use EntraAuth instead of Microsoft.Graph.Authentication, both will be supported for now.
            $script:GraphEntraAuth = $false
            $EntraAuthModuleAvailable = Get-Module EntraAuth -ListAvailable
            $MSGraphAuthAvailable = Get-Module Microsoft.Graph.Authentication -ListAvailable
      
            if ($EntraAuthModuleAvailable) {
                $AuthToken = Get-EntraToken -Service Graph
                if ($AuthToken) {
                    $script:GraphEntraAuth = $true
                    $currentScopes = $AuthToken.Scopes
                    $AuthTokenType = $AuthToken.tokendata.idtyp.replace('app', 'Application').replace('user', 'Delegated')
                }
            }

            #EntraAuth has priority if already connected.
            if ($MSGraphAuthAvailable -and !$script:GraphEntraAuth) {
                $MgGraphContext = Get-MgContext
                $currentScopes = $MgGraphContext.Scopes
                $AuthTokenType = (""+$MgGraphContext.AuthType).replace('AppOnly', 'Application')
            }

            if(!$EntraAuthModuleAvailable -and !$MSGraphAuthAvailable) {
                Write-Warning ("Missing EntraAuth PowerShell module. Please install it with:" + [Environment]::NewLine + "Install-Module EntraAuth") 
                return $false
            }

            if (!($currentScopes)) {
                Write-Warning  ("Not Connected to Microsoft Graph" + `
                        [Environment]::NewLine + "Please connect to Microsoft Graph before running this cmdlet." + `
                        [Environment]::NewLine + "Commercial Tenant: Connect-EntraService -ClientID Graph -Scopes " + ($Scopes -join ",") + `
                        [Environment]::NewLine + "US Gov (GCC-H) Tenant: Connect-EntraService -ClientID Graph " + ($Scopes -join ",") + " -Environment USGov")
                return $false
            }

            if ($AuthType -and $AuthTokenType -ne $AuthType) {
                Write-Warning "Wrong Permission Type: $AuthTokenType, this PowerShell cmdlet requires: $AuthType"
                return $false
            }
            $strScope = ""
            $strAltScope = ""
            $missingScopes = ""
            $missingAltScopes = ""
            $missingScope = $false
            $missingAltScope = $false
            foreach ($scope in $Scopes) {
                $strScope += "`"" + $scope + "`","
                if ($scope -notin $currentScopes) {
                    $missingScope = $true
                    $missingScopes += $scope + ","
                }
            }
            if ($missingScope -and $AltScopes) {
                foreach ($altScope in $AltScopes) {
                    $strAltScope += "`"" + $altScope + "`","
                    if ($altScope -notin $currentScopes) {
                        $missingAltScope = $true
                        $missingAltScopes += $altScope + ","
                    }
                }
            }
            else {
                $missingAltScope = $true
            }
            #If scopes are missing we need to connect using the required scopes
            if ($missingScope -and $missingAltScope) {
                if ($Scopes -and $AltScopes) {
                    Write-Warning  ("Missing scope(s): " + $missingScopes.Substring(0, $missingScopes.Length - 1) + " and missing alternative Scope(s): " + $missingAltScopes.Substring(0, $missingAltScopes.Length - 1) + `
                            [Environment]::NewLine + "Please reconnect to Microsoft Graph before running this cmdlet." + `
                            [Environment]::NewLine + "Commercial Tenant: Connect-EntraService -ClientID Graph -Scopes " + $strScope.Substring(0, $strScope.Length - 1) + " or Connect-EntraService -ClientID Graph -Scopes " + $strAltScope.Substring(0, $strAltScope.Length - 1) + `
                            [Environment]::NewLine + "US Gov (GCC-H) Tenant: Connect-EntraService -ClientID Graph -Environment USGov -Scopes " + $strScope.Substring(0, $strScope.Length - 1) + " or Connect-EntraService -ClientID Graph -Environment USGov -Scopes " + $strAltScope.Substring(0, $strAltScope.Length - 1) )
                }
                else {
                    Write-Warning  ("Missing scope(s): " + $missingScopes.Substring(0, $missingScopes.Length - 1) + `
                            [Environment]::NewLine + "Please reconnect to Microsoft Graph before running this cmdlet." + `
                            [Environment]::NewLine + "Commercial Tenant: Connect-EntraService -ClientID Graph -Scopes " + $strScope.Substring(0, $strScope.Length - 1) + `
                            [Environment]::NewLine + "US Gov (GCC-H) Tenant: Connect-EntraService -ClientID Graph -Scopes " + $strScope.Substring(0, $strScope.Length - 1))
                }
                return $false
            }
            return $true
        }
        "TeamsPowerShell" { 
            #Checking if MicrosoftTeams module is installed
            if (!(Get-Module MicrosoftTeams -ListAvailable)) {
                Write-Warning ("Missing MicrosoftTeams PowerShell module. Please install it with:" + [Environment]::NewLine + "Install-Module MicrosoftTeams") 
                return $false
            }
            #We need to use a cmdlet to know if we are connected to MicrosoftTeams PowerShell
            try {
                Get-CsTenant -ErrorAction SilentlyContinue | Out-Null
                return $true
            }
            catch [System.UnauthorizedAccessException] {
                Write-Warning ("Please connect to Microsoft Teams PowerShell with Connect-MicrosoftTeams before running this cmdlet")
                return $false
            }
        }
        "TeamsDeviceTAC" {
            #Checking if EntraAuth module is installed
            if (!(Get-Module EntraAuth -ListAvailable)) {
                Write-Warning ("Missing EntraAuth PowerShell module. Please install it with:" + [Environment]::NewLine + "Install-Module EntraAuth") 
                return $false
            }
            if (Get-EntraToken TeamsDeviceTAC) {
                return $true
            }
            else {
                Write-Warning "Please connect to Teams TAC API with Connect-UcTeamsDeviceTAC before running this cmdlet"
            }
        }
        Default {
            return $false
        }
    }
}

function Connect-UcTeamsDeviceTAC {
    param(
        [string]$TenantID
    )
    #Create TeamsDeviceTAC Entra Service if doens't exist
    if (!(Get-EntraService -Name TeamsDeviceTAC)) {
        $teamsDeviceTacCfg = @{
            Name          = 'TeamsDeviceTAC'
            ServiceUrl    = 'https://admin.devicemgmt.teams.microsoft.com/'
            Resource      = 'https://admin.devicemgmt.teams.microsoft.com'
            DefaultScopes = @("https://devicemgmt.teams.microsoft.com/.default")
            HelpUrl       = ''
            Header        = @{ }
            NoRefresh     = $false
        }
        Register-EntraService @teamsDeviceTacCfg
    }
    if ($TenantID) {
        Connect-EntraService  -ClientId '12128f48-ec9e-42f0-b203-ea49fb6af367' -Service TeamsDeviceTAC -TenantID $TenantID
    }
    else {
        Connect-EntraService  -ClientId '12128f48-ec9e-42f0-b203-ea49fb6af367' -Service TeamsDeviceTAC
    }
}

function Get-UcArch {
    <#
        .SYNOPSIS
        Funcion to get the Architecture from .exe file

        .DESCRIPTION
        Based on PowerShell script Get-ExecutableType.ps1 by David Wyatt, please check the complete script in:

        Identify 16-bit, 32-bit and 64-bit executables with PowerShell
        https://gallery.technet.microsoft.com/scriptcenter/Identify-16-bit-32-bit-and-522eae75

        .PARAMETER FilePath
        Specifies the executable full file path.

        .EXAMPLE
        PS> Get-UcArch -FilePath C:\temp\example.exe
    #>
    param(
        [string]$FilePath
    )

    try {
        $stream = New-Object System.IO.FileStream(
            $FilePath,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::Read )
        $exeType = 'Unknown'
        $bytes = New-Object byte[](4)
        if ($stream.Seek(0x3C, [System.IO.SeekOrigin]::Begin) -eq 0x3C -and $stream.Read($bytes, 0, 4) -eq 4) {
            if (-not [System.BitConverter]::IsLittleEndian) { [Array]::Reverse($bytes, 0, 4) }
            $peHeaderOffset = [System.BitConverter]::ToUInt32($bytes, 0)

            if ($stream.Length -ge $peHeaderOffset + 6 -and
                $stream.Seek($peHeaderOffset, [System.IO.SeekOrigin]::Begin) -eq $peHeaderOffset -and
                $stream.Read($bytes, 0, 4) -eq 4 -and
                $bytes[0] -eq 0x50 -and $bytes[1] -eq 0x45 -and $bytes[2] -eq 0 -and $bytes[3] -eq 0) {
                $exeType = 'Unknown'
                if ($stream.Read($bytes, 0, 2) -eq 2) {
                    if (-not [System.BitConverter]::IsLittleEndian) { [Array]::Reverse($bytes, 0, 2) }
                    $machineType = [System.BitConverter]::ToUInt16($bytes, 0)
                    switch ($machineType) {
                        0x014C { $exeType = 'x86' }
                        0x8664 { $exeType = 'x64' }
                    }
                }
            }
        }
        return $exeType
    }
    catch {
        return "Unknown"
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Get-UcEntraObjectsOwnedByUser {
    <#
        .SYNOPSIS
        Returns all Entra objects associated with a user.

        .DESCRIPTION
        This function returns the list of Entra Objects associated with the specified user.

        Contributors: Jimmy Vincent, David Paulino

        Requirements:   EntraAuth PowerShell Module (Install-Module EntraAuth)
                        or
                        Microsoft Graph Authentication PowerShell Module (Install-Module Microsoft.Graph.Authentication)

                        Microsoft Graph Scopes:
                            "User.Read.All" or "Directory.Read.All"

        .PARAMETER User
        Specifies the user UPN or User Object ID.

        .PARAMETER Type
        Specifies a filter, valid options:
            Application
            ServicePrincipal
            TokenLifetimePolicy
            SecurityGroup
            DistributionGroup
            Microsoft365Group
            Team
            Yammer

        .EXAMPLE
        PS> Get-UcObjectsOwnedByUser -User user@uclobby.com
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$User,
        [ValidateSet("Application", "ServicePrincipal", "TokenLifetimePolicy", "MailEnabledGroup", "SecurityGroup", "DistributionGroup", "Microsoft365Group", "Team", "Yammer")]
        [string]$Type
    )
    
    #region Graph Connection, Scope validation and module version
    if (!(Test-UcServiceConnection -Type MSGraph -Scopes "User.Read.All" -AltScopes "Directory.Read.All")) {
        return
    }
    #2025-01-31: Only need to check this once per PowerShell session
    if (!($global:UCLobbyTeamsModuleCheck)) {
        Test-UcPowerShellModule -ModuleName UcLobbyTeams | Out-Null
        $global:UCLobbyTeamsModuleCheck = $true
    }
    #endregion

    $output = [System.Collections.ArrayList]::new()
    $graphRequests = [System.Collections.ArrayList]::new()
    $gRequestTmp = [PSCustomObject]@{
        id     = "UserObjects"
        method = "GET"
        url    = "/users/$User/ownedObjects?`$select=id,displayName,visibility,createdDateTime,creationOptions,groupTypes,securityEnabled,mailEnabled"
    }
    [void]$graphRequests.Add($gRequestTmp)
    $BatchResponse = Invoke-UcGraphRequest -Requests $graphRequests -IncludeBody 
        
    if ($BatchResponse.status -eq 404) {
        Write-Warning "User $user was not found, please check it and try again."
        return
    }
        
    $TypeFilter = "All"
    if ($Type) {
        $TypeFilter = $Type
    }
    if ($BatchResponse.body.value.count -gt 225) {
        Write-Warning ("Please be aware that user $user currently has " + $BatchResponse.body.value.count + " Entra Objects, the current limitation is 250 Entra Objects.")
    }
    foreach ($OwnedObject in $BatchResponse.body.value) {
        $tmpType = "NA"
        switch ($OwnedObject.'@odata.type') {
            "#microsoft.graph.application" { 
                if ($TypeFilter -in ("Application" , "All")) {
                    $tmpType = "Application" 
                }
            }
            "#microsoft.graph.servicePrincipal" {
                if ($TypeFilter -in ("ServicePrincipal" , "All")) {
                    $tmpType = "Service Principal"
                }
            }
            "#microsoft.graph.tokenLifetimePolicy" { 
                if ($TypeFilter -in ("TokenLifetimePolicy" , "All")) {
                    $tmpType = "Token Lifetime Policy" 
                }
            }
            "#microsoft.graph.group" { 
                if ($OwnedObject.mailEnabled -and $OwnedObject.securityEnabled -and ($OwnedObject.groupTypes.count -eq 0) -and $TypeFilter -in ("MailEnabledGroup" , "All")) {
                    $tmpType = "Mail enabled security group" 
                }
                if (!($OwnedObject.mailEnabled) -and $OwnedObject.securityEnabled -and ($OwnedObject.groupTypes.count -eq 0) -and $TypeFilter -in ("SecurityGroup" , "All")) {
                    $tmpType = "Security group" 
                }
                if ($OwnedObject.mailEnabled -and !($OwnedObject.securityEnabled) -and ($OwnedObject.groupTypes.count -eq 0) -and $TypeFilter -in ("DistributionGroup" , "All")) {
                    $tmpType = "Distribution group" 
                }
                if ($OwnedObject.groupTypes -contains "Unified" -and $TypeFilter -in ("Microsoft365Group" , "All")) {
                    $tmpType = "Microsoft 365 group" 
                }
                if (($OwnedObject.creationOptions -contains "Team") -and $TypeFilter -in ("Team" , "All")) {
                    $tmpType = "Team" 
                }
                if ($OwnedObject.creationOptions -contains "YammerProvisioning" -and $TypeFilter -in ("Yammer" , "All")) {
                    $tmpType = "Yammer" 
                }
            }
            Default { 
                if ($TypeFilter -eq "All") {
                    $tmpType = $OwnedObject.'@odata.type' 
                }
            }
        }
        if ($tmpType -ne "NA") {
            $UserObject = [PSCustomObject][Ordered]@{
                User            = $User
                ObjectID        = $OwnedObject.id
                DisplayName     = $OwnedObject.displayName
                Type            = $tmpType
                Visibility      = $OwnedObject.visibility
                CreatedDateTime = $OwnedObject.createdDateTime
            }
            $UserObject.PSObject.TypeNames.Insert(0, 'EntraObjectsOwnedByUser')
            [void]$output.Add($UserObject)
        }
    }
    return $output | Sort-Object Type, DisplayName
}

function Get-UcTeamsDevice {
    <#
        .SYNOPSIS
        Get Microsoft Teams Devices information

        .DESCRIPTION
        This function fetch Teams Devices provisioned in a M365 Tenant using MS Graph.

        Contributors: David Paulino, Silvio Schanz, Gonçalo Sepulveda, Bryan Kendrick, Daniel Jelinek and Traci Herr

        Requirements:   EntraAuth PowerShell Module (Install-Module EntraAuth)
                        or
                        Microsoft Graph Authentication PowerShell Module (Install-Module Microsoft.Graph.Authentication)

                        Microsoft Graph Scopes:
                                "TeamworkDevice.Read.All"
                                "User.Read.All"

                        Note: UseTac parameter requires EntraAuth PowerShell Module

        .PARAMETER TACDeviceID
        Allows specifying a single Teams Device using the Device ID used by TAC.
        
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
        When present will export the detailed results to a CSV file. By default, will save the file under the current user downloads, unless we specify the OutputPath.

        .PARAMETER OutputPath
        Allows to specify the path where we want to save the results.

        .PARAMETER UseTAC
        When present it will use the Teams Admin Center API to get the Teams Devices information.

        .EXAMPLE 
        PS> Get-UcTeamsDevice

        .EXAMPLE 
        PS> Get-UcTeamsDevice -Filter MTR

        .EXAMPLE
        PS> Get-UcTeamsDevice -Detailed

    #>
    [cmdletbinding()]
    param(
        [string]$TACDeviceID,
        [ValidateSet("Phone", "MTR", "MTRA", "MTRW", "SurfaceHub", "Display", "Panel", "SIPPhone")]
        [string]$Filter,
        [switch]$Detailed,
        [switch]$ExportCSV,
        [string]$OutputPath,
        [switch]$UseTAC
    )

    $outTeamsDevices = [System.Collections.ArrayList]::new()
    $devicesProcessed = 0
    $BaseDevicesAPIPath = "api/v2/devices"

    if ($ExportCSV) {
        $Detailed = $true
    }

    #2025-01-31: Only need to check this once per PowerShell session
    if (!($global:UCLobbyTeamsModuleCheck)) {
        Test-UcPowerShellModule -ModuleName UcLobbyTeams | Out-Null
        $global:UCLobbyTeamsModuleCheck = $true
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
    
    if ($UseTAC) {
        #region 2025-01-29: Using TAC API
        if (Test-UcServiceConnection -Type TeamsDeviceTAC) {
            if ($TACDeviceID) {
                try { 
                    $TeamsDevices = Invoke-EntraRequest -Path ($BaseDevicesAPIPath + "/" + $TACDeviceID) -Service TeamsDeviceTAC
                }
                catch {
                    Write-Warning "Please check the TAC Device Id ($TACDeviceID) and try again."
                    return $null
                }
            }
            else {
                $tmpFileName = "MSTeamsDevicesTAC_" + $Filter + "_" + ( get-date ).ToString('yyyyMMdd-HHmmss') + ".csv"
                switch ($Filter) {
                    "Phone" { $DeviceFilter = "{`"deviceTypes`":{`"eq`":[`"IpPhone`",`"LowCostPhone`"]}}" }
                    "MTR" { $DeviceFilter = "{`"deviceTypes`":{`"eq`":[`"teamsRoom`",`"collaborationBar`",`"touchConsole`"]}}" }
                    "MTRW" { $DeviceFilter = "{`"deviceTypes`":{`"eq`":[`"teamsRoom`"]}}" }
                    "MTRA" { $DeviceFilter = "{`"deviceTypes`":{`"eq`":[`"collaborationBar`",`"touchConsole`"]}}" }
                    "SurfaceHub" { $DeviceFilter = "{`"deviceTypes`":{`"eq`":[`"surfaceHub`"]}}" }
                    "Display" { $DeviceFilter = "{`"deviceTypes`":{`"eq`":[`"teamsDisplay`"]}}" }
                    "Panel" { $DeviceFilter = "{`"deviceTypes`":{`"eq`":[`"teamsPanel`"]}}" }
                    "SIPPhone" { $DeviceFilter = "{`"deviceTypes`":{`"eq`":[`"sip`"]}}" }
                    default {
                        $DeviceFilter = $null
                        $tmpFileName = "MSTeamsDevices_All_" + ( get-date ).ToString('yyyyMMdd-HHmmss') + ".csv"
                    }
                }
                if ($DeviceFilter) {
                    $RequestPath = $BaseDevicesAPIPath + "?limit=50&filterJson=" + [System.Web.HttpUtility]::UrlEncode($DeviceFilter) + "&fetchCurrentSoftwareVersions=true"
                }
                else {
                    $RequestPath = $BaseDevicesAPIPath + "?limit=50&fetchCurrentSoftwareVersions=true"
                }
                
                #region 2025-09-25: Adding page iterarion.
                Write-Progress -Activity "Getting Teams Device Info from TAC" -Status "0 of ?"
                $TACReply = Invoke-EntraRequest -Path $RequestPath -Service TeamsDeviceTAC
                $TeamsDevices = $TACReply.devices
                $TACDeviceCount = $TACReply.devices.count
                Write-Verbose "First Request with $TACDeviceCount devices."
                while(!([string]::IsNullOrEmpty($TACReply.continuationToken))){
                    Write-Progress -Activity "Getting Teams Device Info from TAC" -Status "$TACDeviceCount of ?"
                    $RequestNextPage = $RequestPath+"&continuationToken=" + [System.Web.HttpUtility]::UrlEncode($TACReply.continuationToken)
                    Write-Verbose "Requesting next page with continuation token: $RequestNextPage"
                    $TACReply = Invoke-EntraRequest -Path $RequestNextPage -Service TeamsDeviceTAC
                    $TeamsDevices += $TACReply.devices
                    $TACDeviceCount += $TACReply.devices.count
                }
                #endregion
            }
            Write-Verbose ("Finished fetching devices, found $TACDeviceCount devices in TAC.")
            foreach ($TeamsDevice in $TeamsDevices) {
                Write-Progress -Activity "Processing Teams Device information" -Status "$devicesProcessed of $TACDeviceCount"
                $outMacAddress = ""
                $lastTACHeartBeat = $null
                foreach ($macAddressInfo in $TeamsDevice.macAddressInfos) {
                    $outMacAddress += $macAddressInfo.interfaceType + ":" + $macAddressInfo.macAddress + ";"
                }

                #We need to use SignInState timestamp or if not available last health status change.
                if ($TeamsDevice.signInInfo.signInState) {
                    $lastTACHeartBeat = (Get-Date "01/01/1970").AddMilliseconds($TeamsDevice.signInInfo.timestamp).ToLocalTime()
                }
                else {
                    $lastTACHeartBeat = (Get-Date "01/01/1970").AddSeconds($TeamsDevice.healthSummary.since).ToLocalTime()
                }

                if ($Detailed) {
                    #Getting the Health and Operations (commands) Information
                    #$RequestHealthPath = $BaseDevicesAPIPath + "/" + $TeamsDevice.baseInfo.id + "/health"
                    #$TeamsDeviceHealth = Invoke-EntraRequest -Path $RequestHealthPath -Service TeamsDeviceTAC
                    try{
                        $RequestOperationPath = $BaseDevicesAPIPath + "/" + $TeamsDevice.baseInfo.id + "/commands?FetchInitiatorInfo=true"
                        $LastTeamsDeviceOperation = (Invoke-EntraRequest -Path $RequestOperationPath -Service TeamsDeviceTAC -ErrorAction SilentlyContinue).commands | Sort-Object queuedAt -Descending | Select-Object -First 1
                    } catch{}

                    $lastHistoryInitiatedByName = "System (Automatic)"
                    $lastHistoryInitiatedByUpn = ""
                    if ($LastTeamsDeviceOperation.initiatedBy) {
                        $lastHistoryInitiatedByName = $LastTeamsDeviceOperation.initiatedBy.userName
                        $lastHistoryInitiatedByUpn = $LastTeamsDeviceOperation.initiatedBy.upn
                    }

                    $TDObj = [PSCustomObject][Ordered]@{
                        TACDeviceID                = $TeamsDevice.baseInfo.id
                        DeviceType                 = Convert-UcTeamsDeviceType $TeamsDevice.deviceType
                        Manufacturer               = $TeamsDevice.deviceModelRef.manufacturer
                        Model                      = $TeamsDevice.deviceModelRef.model
                        UserDisplayName            = $TeamsDevice.lastLoggedInUserRef.userName
                        UserUPN                    = $TeamsDevice.lastLoggedInUserRef.upn
                        License                    = $TeamsDevice.licenseDetails.effectiveLicense.friendlyName
                        LastLicenseUpdate          = (Get-Date "01/01/1970").AddSeconds($TeamsDevice.licenseDetails.updatedAt).ToLocalTime()
                        SignInMode                 = Convert-UcTeamsDeviceSignInMode $TeamsDevice.userType
                        SignInState                = $TeamsDevice.signInInfo.signInState
                        DeviceHealth               = $TeamsDevice.healthSummary.healthState
                        LastTACHeartBeat           = $lastTACHeartBeat
                        Notes                      = $TeamsDevice.notes
                        CompanyAssetTag            = $TeamsDevice.companyAssetTag
                        SerialNumber               = $TeamsDevice.deviceIds.oemSerialNumber 
                        MacAddresses               = $outMacAddress
                        ipAddress                  = $TeamsDevice.ipAddress
                        PairingStatus              = $TeamsDevice.pairingStatus
                            
                        #Current Versions
                        TeamsAdminAgentVersion     = $TeamsDevice.softwareVersions.adminagent.versionname
                        FirmwareVersion            = $TeamsDevice.softwareVersions.firmware.versionName
                        CompanyPortalVersion       = $TeamsDevice.softwareVersions.companyPortal.versionName 
                        OEMAgentAppVersion         = $TeamsDevice.softwareVersions.partnerAgent.versionName 
                        TeamsAppVersion            = $TeamsDevice.softwareVersions.teamsApp.versionName
                        AuthenticatorVersion       = $TeamsDevice.softwareVersions.authenticatorApp.versionName
                        MicrosoftIntuneVersion     = $TeamsDevice.softwareVersions.microsoftIntuneApp.versionName
                        
                        AutomaticUpdates           = Convert-UcTeamsDeviceAutoUpdateDays $TeamsDevice.administrationConfig.autoUpdateFrequencyInDays
                        CurrentFirmwareRing        = $TeamsDevice.currentFirmwareRing
                        PreviewBuilds              = $TeamsDevice.softwareConfig.isOptedForPreviewBuilds
                        ConfigurationProfileId     = $TeamsDevice.configRef.configId
                        ConfigurationProfileName   = $TeamsDevice.configRef.name

                        LastHistoryAction          = $LastTeamsDeviceOperation.command
                        LastHistoryStatus          = $LastTeamsDeviceOperation.commandStatus
                        LastHistoryInitiatedByName = $lastHistoryInitiatedByName
                        LastHistoryInitiatedByUpn  = $lastHistoryInitiatedByUpn
                        LastHistoryModifiedDate    = (Get-Date "01/01/1970").AddSeconds($LastTeamsDeviceOperation.baseInfo.modifiedAt).ToLocalTime()

                        WhenCreated                = (Get-Date "01/01/1970").AddSeconds($TeamsDevice.baseInfo.createdAt).ToLocalTime()
                        WhenChanged                = (Get-Date "01/01/1970").AddSeconds($TeamsDevice.baseInfo.modifiedAt).ToLocalTime()
                    }
                }
                else {
                    $TDObj = [PSCustomObject][Ordered]@{
                        TACDeviceID              = $TeamsDevice.baseInfo.id
                        DeviceType               = Convert-UcTeamsDeviceType $TeamsDevice.deviceType
                        Manufacturer             = $TeamsDevice.deviceModelRef.manufacturer
                        Model                    = $TeamsDevice.deviceModelRef.model
                        UserDisplayName          = $TeamsDevice.lastLoggedInUserRef.userName
                        UserUPN                  = $TeamsDevice.lastLoggedInUserRef.upn
                        License                  = $TeamsDevice.licenseDetails.effectiveLicense.friendlyName
                        LastLicenseUpdate        = (Get-Date "01/01/1970").AddSeconds($TeamsDevice.licenseDetails.updatedAt).ToLocalTime()
                        SignInMode               = Convert-UcTeamsDeviceSignInMode $TeamsDevice.userType
                        SignInState              = $TeamsDevice.signInInfo.signInState
                        DeviceHealth             = $TeamsDevice.healthSummary.healthState
                        LastTACHeartBeat         = $lastTACHeartBeat
                        Notes                    = $TeamsDevice.notes
                        CompanyAssetTag          = $TeamsDevice.companyAssetTag
                        SerialNumber             = $TeamsDevice.deviceIds.oemSerialNumber 
                        MacAddresses             = $outMacAddress
                        ipAddress                = $TeamsDevice.ipAddress
                        PairingStatus            = $TeamsDevice.pairingStatus

                        #Current Versions
                        TeamsAdminAgentVersion   = $TeamsDevice.softwareVersions.adminagent.versionname
                        FirmwareVersion          = $TeamsDevice.softwareVersions.firmware.versionName
                        CompanyPortalVersion     = $TeamsDevice.softwareVersions.companyPortal.versionName 
                        OEMAgentAppVersion       = $TeamsDevice.softwareVersions.partnerAgent.versionName 
                        TeamsAppVersion          = $TeamsDevice.softwareVersions.teamsApp.versionName
                        AuthenticatorVersion     = $TeamsDevice.softwareVersions.authenticatorApp.versionName
                        MicrosoftIntuneVersion   = $TeamsDevice.softwareVersions.microsoftIntuneApp.versionName                        
                                                   
                        AutomaticUpdates         = Convert-UcTeamsDeviceAutoUpdateDays $TeamsDevice.administrationConfig.autoUpdateFrequencyInDays
                        CurrentFirmwareRing      = $TeamsDevice.currentFirmwareRing
                        PreviewBuilds            = $TeamsDevice.softwareConfig.isOptedForPreviewBuilds
                        ConfigurationProfileId   = $TeamsDevice.configRef.configId
                        ConfigurationProfileName = $TeamsDevice.configRef.name
                        WhenCreated              = (Get-Date "01/01/1970").AddSeconds($TeamsDevice.baseInfo.createdAt).ToLocalTime()
                        WhenChanged              = (Get-Date "01/01/1970").AddSeconds($TeamsDevice.baseInfo.modifiedAt).ToLocalTime()
                    }
                    $TDObj.PSObject.TypeNames.Insert(0, 'TeamsDevice')
                }
                [void]$outTeamsDevices.Add($TDObj)
                $devicesProcessed++
            }
        }
        #endregion
    }
    else {
        if (!(Test-UcServiceConnection -Type MSGraph -Scopes "TeamworkDevice.Read.All", "User.Read.All" -AltScopes ("TeamworkDevice.Read.All", "Directory.Read.All"))) {
            return
        }

        $graphRequests = [System.Collections.ArrayList]::new()
        $tmpFileName = "MSTeamsDevices_" + $Filter + "_" + ( get-date ).ToString('yyyyMMdd-HHmmss') + ".csv"
        if ($TACDeviceID) {
            $gRequestTmp = [PSCustomObject]@{
                id     = "ipPhone"
                method = "GET"
                url    = "/teamwork/devices/" + $TACDeviceID 
            }
            [void]$graphRequests.Add($gRequestTmp)
        }
        else {
            switch ($filter) {
                "Phone" { 
                    $gRequestTmp = [PSCustomObject]@{
                        id     = "ipPhone"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'ipPhone'"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                    $gRequestTmp = [PSCustomObject]@{
                        id     = "lowCostPhone"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'lowCostPhone'"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                }
                "MTR" {
                    $gRequestTmp = [PSCustomObject]@{
                        id     = "teamsRoom"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'teamsRoom'"
                    }
                    [void]$graphRequests.Add($gRequestTmp) 
                    $gRequestTmp = [PSCustomObject]@{
                        id     = "collaborationBar"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'collaborationBar'"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                    $gRequestTmp = [PSCustomObject]@{
                        id     = "touchConsole"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'touchConsole'"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                }
                "MTRW" {
                    $gRequestTmp = [PSCustomObject]@{
                        id     = "teamsRoom"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'teamsRoom'"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                }
                "MTRA" {            
                    $gRequestTmp = [PSCustomObject]@{
                        id     = "collaborationBar"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'collaborationBar'"
                    }
                    [void]$graphRequests.Add($gRequestTmp) 
                    $gRequestTmp = [PSCustomObject]@{
                        id     = "touchConsole"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'touchConsole'"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                }
                "SurfaceHub" {
                    $gRequestTmp = [PSCustomObject]@{
                        id     = "surfaceHub"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'surfaceHub'"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                }
                "Display" {
                    $gRequestTmp = [PSCustomObject]@{
                        id     = "teamsDisplay"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'teamsDisplay'"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                }
                "Panel" {
                    $gRequestTmp = [PSCustomObject]@{
                        id     = "teamsPanel"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'teamsPanel'"
                    }
                    [void]$graphRequests.Add($gRequestTmp) 
                }
                "SIPPhone" {
                    $gRequestTmp = [PSCustomObject]@{
                        id     = "sip"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'sip'"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                }
                default {
                    $gRequestTmp = [PSCustomObject]@{
                        id     = 1
                        method = "GET"
                        url    = "/teamwork/devices"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                    $tmpFileName = "MSTeamsDevices_All_" + ( get-date ).ToString('yyyyMMdd-HHmmss') + ".csv"
                }
            }
        }
        $TeamsDeviceList = (Invoke-UcGraphRequest -Requests $graphRequests -Beta -Activity "Get-UcTeamsDevice, getting Teams device info").value
        
        #To improve performance we will use batch requests
        $graphRequests = [System.Collections.ArrayList]::new()
        foreach ($TeamsDevice in $TeamsDeviceList) {
            if (($graphRequests.id -notcontains $TeamsDevice.currentuser.id) -and !([string]::IsNullOrEmpty($TeamsDevice.currentuser.id))) {
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = $TeamsDevice.currentuser.id
                    method = "GET"
                    url    = "/users/" + $TeamsDevice.currentuser.id
                }
                [void]$graphRequests.Add($gRequestTmp)
            }
            if ($Detailed) {
                $gRequestTmp = [PSCustomObject]@{
                    id     = $TeamsDevice.id + "-activity"
                    method = "GET"
                    url    = "/teamwork/devices/" + $TeamsDevice.id + "/activity"
                }
                [void]$graphRequests.Add($gRequestTmp)
                $gRequestTmp = [PSCustomObject]@{
                    id     = $TeamsDevice.id + "-configuration"
                    method = "GET"
                    url    = "/teamwork/devices/" + $TeamsDevice.id + "/configuration"
                }
                [void]$graphRequests.Add($gRequestTmp)
                $gRequestTmp = [PSCustomObject]@{
                    id     = $TeamsDevice.id + "-health"
                    method = "GET"
                    url    = "/teamwork/devices/" + $TeamsDevice.id + "/health"
                }
                [void]$graphRequests.Add($gRequestTmp)
                $gRequestTmp = [PSCustomObject]@{
                    id     = $TeamsDevice.id + "-operations"
                    method = "GET"
                    url    = "/teamwork/devices/" + $TeamsDevice.id + "/operations"
                }
                [void]$graphRequests.Add($gRequestTmp)
            } 
        }
        if ($graphRequests.Count -gt 0) {
            if ($Detailed) {
                $ActivityInfo = "Get-UcTeamsDevice, getting Teams device addtional information (User UPN/Health/Operations/Configurarion)."
            }
            else {
                $ActivityInfo = "Get-UcTeamsDevice, getting Teams device user information."
            }
            $graphResponseExtra = (Invoke-UcGraphRequest -Requests $graphRequests -Beta -Activity $ActivityInfo -IncludeBody)
        }
            
        foreach ($TeamsDevice in $TeamsDeviceList) {
            $devicesProcessed++
            $userUPN = ($graphResponseExtra | Where-Object { $_.id -eq $TeamsDevice.currentuser.id }).body.userPrincipalName
            $tmpDeviceType = ""
            if ($TeamsDevice.deviceType) {
                $tmpDeviceType = Convert-UcTeamsDeviceType $TeamsDevice.deviceType
            }

            $outMacAddress = ""
            foreach ($macAddress in $TeamsDevice.hardwaredetail.macAddresses) {
                $outMacAddress += $macAddress + ";"
            }

            #region 2025-05-11: Fix improving date parsing
            $tmpWhenCreated = ""
            if ($TeamsDevice.createdDateTime) {
                $tmpWhenCreated = (Get-Date ($TeamsDevice.createdDateTime)).ToLocalTime()
            }
            if ($TeamsDevice.lastModifiedDateTime) {
                $tmpWhenChanged = (Get-Date ($TeamsDevice.lastModifiedDateTime)).ToLocalTime()
            }
            #endregion

            if ($Detailed) {
                $TeamsDeviceConfiguration = ""
                $TeamsDeviceActivity = ($graphResponseExtra | Where-Object { $_.id -eq ($TeamsDevice.id + "-activity") }).body
                $TeamsDeviceConfiguration = ($graphResponseExtra | Where-Object { $_.id -eq ($TeamsDevice.id + "-configuration") }).body
                $TeamsDeviceHealth = ($graphResponseExtra | Where-Object { $_.id -eq ($TeamsDevice.id + "-health") }).body
                $TeamsDeviceOperations = ($graphResponseExtra | Where-Object { $_.id -eq ($TeamsDevice.id + "-operations") }).body.value

                if ($TeamsDeviceOperations.count -gt 0) {
                    $LastHistoryAction = $TeamsDeviceOperations[0].operationType
                    $LastHistoryStatus = $TeamsDeviceOperations[0].status
                    $LastHistoryInitiatedBy = $TeamsDeviceOperations[0].createdBy.user.displayName
                    $LastHistoryModifiedDate = ""
                    if(($TeamsDeviceOperations[0].lastActionDateTime)){
                        $LastHistoryModifiedDate = (Get-Date $TeamsDeviceOperations[0].lastActionDateTime).ToLocalTime()
                    }
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

                #region 2025-03-31: Fix "You cannot call a method on a null-valued expression" if date were empty.
                $tmpConfigurationCreateDate = ""
                $tmpConfigurationLastModifiedDate = ""
                if ($TeamsDeviceConfiguration.createdDateTime) {
                    $tmpConfigurationCreateDate = (Get-Date $TeamsDeviceConfiguration.createdDateTime).ToLocalTime()
                }
                if ($TeamsDeviceConfiguration.lastModifiedDateTime) {
                    $tmpConfigurationLastModifiedDate = (Get-Date $TeamsDeviceConfiguration.lastModifiedDateTime).ToLocalTime()
                }
                #endregion

                $tmpConnectionLastActivity = ""
                if ($TeamsDeviceHealth.connection.lastModifiedDateTime){
                    $tmpConnectionLastActivity = (Get-Date $TeamsDeviceHealth.connection.lastModifiedDateTime).ToLocalTime()
                }
               
                $TDObj = [PSCustomObject][Ordered]@{
                    TACDeviceID                   = $TeamsDevice.id
                    DeviceType                    = $tmpDeviceType
                    Manufacturer                  = $TeamsDevice.hardwaredetail.manufacturer
                    Model                         = $TeamsDevice.hardwaredetail.model
                    UserDisplayName               = $TeamsDevice.currentuser.displayName
                    UserUPN                       = $userUPN 
                    Notes                         = $TeamsDevice.notes
                    CompanyAssetTag               = $TeamsDevice.companyAssetTag
                    SerialNumber                  = $TeamsDevice.hardwaredetail.serialNumber 
                    MacAddresses                  = $outMacAddress
                    DeviceHealth                  = $TeamsDevice.healthStatus
                    WhenCreated                   = $tmpWhenCreated
                    WhenChanged                   = $tmpWhenChanged
                    ChangedByUser                 = $TeamsDevice.lastModifiedBy.user.displayName
        
                    #Activity
                    ActivePeripherals             = $TeamsDeviceActivity.activePeripherals
        
                    #Configuration
                    ConfigurationCreateDate       = $tmpConfigurationCreateDate
                    ConfigurationCreatedBy        = $TeamsDeviceConfiguration.createdBy
                    ConfigurationLastModifiedDate = $tmpConfigurationLastModifiedDate
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
                    #2024-04-17: Added connection fields
                    ConnectionStatus              = $TeamsDeviceHealth.connection.connectionStatus
                    ConnectionLastActivity        = $tmpConnectionLastActivity
                    
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
            }
            else {
                $TDObj = [PSCustomObject][Ordered]@{
                    TACDeviceID     = $TeamsDevice.id
                    DeviceType      = $tmpDeviceType
                    UserDisplayName = $TeamsDevice.currentuser.displayName
                    UserUPN         = $userUPN 
                    Manufacturer    = $TeamsDevice.hardwaredetail.manufacturer
                    Model           = $TeamsDevice.hardwaredetail.model
                    DeviceHealth    = $TeamsDevice.healthStatus
                    #2024-04-19: Adding additional fields that are available on graph api
                    #region Details that are in the device info but only shown if we do Format-List (FL)
                    SerialNumber    = $TeamsDevice.hardwaredetail.serialNumber 
                    MacAddresses    = $outMacAddress
                    WhenCreated     = $tmpWhenCreated
                    WhenChanged     = $tmpWhenChanged
                    ChangedByUser   = $TeamsDevice.lastModifiedBy.user.displayName
                    #endregion
                }
                $TDObj.PSObject.TypeNames.Insert(0, 'TeamsDevice')
            }
            [void]$outTeamsDevices.Add($TDObj)
        }
    }
    #2023-10-20: We only need to output if we have results.
    if ($devicesProcessed -gt 0) {
        #region: Modified by Daniel Jelinek
        if ($ExportCSV) {
            $OutputFullPath = [System.IO.Path]::Combine($OutputPath, $tmpFileName)
            $outTeamsDevices | Sort-Object DeviceType, Manufacturer, Model | Export-Csv -path $OutputFullPath -NoTypeInformation
            Write-Host ("Results available in: " + $OutputFullPath) -ForegroundColor Cyan
        }
        else {
            return $outTeamsDevices | Sort-Object DeviceType, Manufacturer, Model
        }
        #endregion
    }
}

function Get-UcTeamsDeviceConfigurationProfile {
    <#
        .SYNOPSIS
        Returns all Teams Device Configuration Profiles

        .DESCRIPTION
        This function fetch Teams Devices Configuration Profiles using the TAC API.

        Requirements:   EntraAuth PowerShell Module (Install-Module EntraAuth)

        .PARAMETER Filter
        Specifies a filter, valid options:
            Phone - Teams Native Phones
            MTR - Microsoft Teams Rooms running Android
            Display - Microsoft Teams Displays 
            Panel - Microsoft Teams Panels

        .EXAMPLE 
        PS> Get-UcTeamsDeviceConfigurationProfile

        .EXAMPLE 
        PS> Get-UcTeamsDeviceConfigurationProfile -Filter MTR
    #>
    param (
        [String]$Identity,
        [ValidateSet("Phone", "MTR", "Display", "Panel")]
        [string]$Filter
    )
    $outTeamsDeviceConfiguration = [System.Collections.ArrayList]::new()
    $BaseAPIPath = "api/v2/configProfiles"

    if (!($global:UCLobbyTeamsModuleCheck)) {
        Test-UcPowerShellModule -ModuleName UcLobbyTeams | Out-Null
        $global:UCLobbyTeamsModuleCheck = $true
    }

    if (Test-UcServiceConnection -Type TeamsDeviceTAC) {
        if ($Identity) {
            $RequestPath = $BaseAPIPath + "/" + $Identity
            try{
                $ConfigurationProfiles = Invoke-EntraRequest -Path $RequestPath -Service TeamsDeviceTAC
            } catch{ }
        }
        else {
            switch ($Filter) {
                "Phone" { $DeviceFilter = "{`"deviceTypes`":{`"eq`":[`"IpPhone`"]}}" }
                "MTR" { $DeviceFilter = "{`"deviceTypes`":{`"eq`":[`"teamsRoom`",`"collaborationBar`"]}}" }
                "Display" { $DeviceFilter = "{`"deviceTypes`":{`"eq`":[`"teamsDisplay`"]}}" }
                "Panel" { $DeviceFilter = "{`"deviceTypes`":{`"eq`":[`"teamsPanel`"]}}" }
                default {
                    $DeviceFilter = $null
                }
            }
            if ($DeviceFilter) {
                $RequestPath = $BaseAPIPath + "?filterJson=" + [System.Web.HttpUtility]::UrlEncode($DeviceFilter)
            }
            else {
                $RequestPath = $BaseAPIPath
            }
            #TODO: Adding support for Pages, currently we are not expecting more than 50 Configuration Profiles.
            try{
            $ConfigurationProfiles = (Invoke-EntraRequest $RequestPath -Service TeamsDeviceTAC).paginatedConfigProfiles
            } catch {}
        }
        foreach ($ConfigurationProfile in $ConfigurationProfiles) {
            #Maintenance Window
            $tmpMaintenanceWindowDaysOfWeek = ""
            $tmpMaintenanceWindowStartTime = ""
            $tmpMaintenanceWindowDuration = ""
            foreach ($maintenanceSchedule in $ConfigurationProfile.maintenanceConfig.weeklyMaintenanceWindows) {
                $tmpMaintenanceWindowDaysOfWeek += $maintenanceSchedule.dayOfWeek + ","
                $tmpMaintenanceWindowStartTime = $maintenanceSchedule.startTime
                $tmpMaintenanceWindowDuration = $maintenanceSchedule.duration
            }

            #Restart Teams App has 3 Options: Never, When Needed, Daily - HH:MM
            $tmpRestartTeamsApp = ""
            if (![string]::IsNullOrEmpty($ConfigurationProfile.deviceAppRestartConfig.isDeviceAppRestartEnabled)) {
                $tmpRestartTeamsApp = "Never"
                if ($ConfigurationProfile.deviceAppRestartConfig.isDeviceAppRestartEnabled) {
                    $tmpRestartTeamsApp = "When needed"
                    if (!$ConfigurationProfile.deviceAppRestartConfig.isDeviceAppAutoRestartEnabled) {
                        $tmpRestartTimeMinutes = $ConfigurationProfile.deviceAppRestartConfig.scheduledDeviceAppRestartTime
                        if ([string]::IsNullOrEmpty($tmpRestartTimeMinutes)) {
                            $tmpRestartTimeMinutes = 1080
                        }
                        $tmpRestartTeamsApp = "Daily - " + (New-TimeSpan -Minutes $tmpRestartTimeMinutes).ToString("hh\:mm")
                    }
                }
            }
        
            $TDObj = [PSCustomObject][Ordered]@{
                Identity                               = $ConfigurationProfile.baseInfo.id
                DisplayName                            = $ConfigurationProfile.identity
                Description                            = $ConfigurationProfile.description
                DeviceType                             = Convert-UcTeamsDeviceType $ConfigurationProfile.deviceType
            
                #region General Settings
                #All Devices
                DeviceLock                             = $ConfigurationProfile.deviceLock
                DeviceLockTimeoutinSeconds             = $ConfigurationProfile.deviceLockTimeout
                DeviceLockPin                          = $ConfigurationProfile.deviceLockPin

                #MTRoA Only
                HDMIContentSharing                     = $ConfigurationProfile.hdmiIngestConfig.isContentSharingEnabled
                HDMIContentSharingIncludeAudio         = $ConfigurationProfile.hdmiIngestConfig.isAudioSharingEnabled
                HDMIContentSharingAutoSharing          = $ConfigurationProfile.hdmiIngestConfig.isAutoSharingEnabled

                #Phone Only
                EnforceDeviceLock                      = $ConfigurationProfile.isForcePinChangeEnabled
            
                #All Devices
                Language                               = $ConfigurationProfile.language
                Timezone                               = $ConfigurationProfile.timezone
                DateFormat                             = $ConfigurationProfile.dateformat
                TimeFormat                             = $ConfigurationProfile.timeFormat
                MaintenanceWindowDaysOfWeek            = $tmpMaintenanceWindowDaysOfWeek -replace ".$"
                MaintenanceWindowStartTime             = $tmpMaintenanceWindowStartTime -replace ".{3}$"
                MaintenanceWindowDurationInMinutes     = $tmpMaintenanceWindowDuration
                DailyDeviceRestart                     = $ConfigurationProfile.maintenanceConfig.isAutoDeviceRestartEnabled
                DailyDeviceRestartStartTime            = $ConfigurationProfile.maintenanceConfig.dailyMaintenanceWindow.startTime 
                DailyDeviceRestartDurationInMinutes    = $ConfigurationProfile.maintenanceConfig.dailyMaintenanceWindow.duration
            
                #MTRoA Only
                EnableTouchScreenControls              = $ConfigurationProfile.isTouchscreenControlsEnabled
                DualDisplayMode                        = $ConfigurationProfile.dualDisplayModeConfig.isDualDisplayModeEnabled
                DualDisplaySwapScreens                 = $ConfigurationProfile.dualDisplayModeConfig.isSwapScreensEnabled
                LogsAndFeedbackEmail                   = $ConfigurationProfile.logsAndFeedbackConfig.emailAddress
                LogsIncludeFeedback                    = $ConfigurationProfile.logsAndFeedbackConfig.isSendLogsAndFeedbackEnabled

                #Panel Only
                ShowRoomEquipment                      = $ConfigurationProfile.isRoomEquipmentEnabled

                #Phone and Display
                RestartTeamsApp                        = $tmpRestartTeamsApp
                #endregion

                #region Meeting Settings
                #Panel Only
                AllowReservationFromPanel              = $ConfigurationProfile.areRoomReservationsEnabled
                ShowMeetingNames                       = !$ConfigurationProfile.hideMeetingNames

                #Panel and Display (for Display it's under General Settings)
                ShowQRCodeForReservations              = $ConfigurationProfile.isQRCodeForSignInOrReservationEnabled

                #MTRoA and Panel
                RoomCapacityNotification               = $ConfigurationProfile.roomCapacityNotificationEnabled
                AllowReservationExtension              = $ConfigurationProfile.extendRoomReservationEnabled
            
                #MTRoA Only
                AllowWhiteBoard                        = $ConfigurationProfile.isWhiteboardEnabled
                RequirePasscodeToJoin                  = $ConfigurationProfile.isRequirePasscodeForMeetingsEnabled

                #Panel Only
                SendCheckinNotification                = $ConfigurationProfile.checkinNotification
                ReleaseRoomIfNoCheckin                 = $ConfigurationProfile.checkinRoomReleaseConfig.enabled
                ReleaseRoomIfNoCheckinTimeoutInSeconds = $ConfigurationProfile.checkinRoomReleaseConfig.timeout
                EarlyCheckout                          = $ConfigurationProfile.checkoutRoomEnabled
                #endregion

                #region Calling Settings
                #Phone Only
                AdvancedCalling                        = $ConfigurationProfile.isIPPhonePremiumCapSkuEnabled
                CallQualitySurvey                      = $ConfigurationProfile.isCallQualityFeedbackSurveyEnabled
                DisplayCallForwardingOnHomeScreen      = $ConfigurationProfile.isIPPhoneCallForwardingOnTheMainScreenEnabled 
            
                #Phone and Display, on display it's the Virtual FrontDesk
                Hotline                                = $ConfigurationProfile.hotlineConfig.hotlineIsEnabled
                HotlineDisplayName                     = $ConfigurationProfile.hotlineConfig.hotlineDisplayName

                #Display Only
                VirtualFrontDeskVideoEnabled           = $ConfigurationProfile.hotlineConfig.isVideoDefaultOn
                VirtualFrontDeskWelcomeMessage         = $ConfigurationProfile.hotlineConfig.welcomeMessage
                #endregion

                #region Device Settings
                #Panel Only
                BusyLightColor                         = $ConfigurationProfile.usageStateIndicatorBusyStateColor.name
                #MTRoA and Panel
                Background                             = $ConfigurationProfile.theme.name

                #All Devices
                DisplayScreenSaver                     = $ConfigurationProfile.displayScreenSaver
                DisplayScreenSaverTimeoutInSeconds     = $ConfigurationProfile.screenSaverTimeout

                #All Devices
                DisplayBacklightBrightness             = $ConfigurationProfile.displayBacklitBrightness
                DisplayBacklightTimeoutInSeconds       = $ConfigurationProfile.displayBacklitTimeout
                DisplayHighContrast                    = $ConfigurationProfile.displayHighContrast
                SilentMode                             = $ConfigurationProfile.silentMode
                OfficeStartHours                       = $ConfigurationProfile.officeStartHours
                OfficeEndHours                         = $ConfigurationProfile.officeEndHours
                PowerSaving                            = $ConfigurationProfile.powerSaving
                ScreenCapture                          = $ConfigurationProfile.screenCapture
            
                #MTRoA Only
                BluetoothBeaconing                     = $ConfigurationProfile.bluetoothBeconingEnabled
                BtBeaconAcceptProximityInvitations     = $ConfigurationProfile.roomQrCodeConfig.isAcceptProximityInvitationEnabled
                RemoteControlFromPersonalDevices       = $ConfigurationProfile.allowRoomRemoteEnabled
                ShowRoomQRcode                         = $ConfigurationProfile.roomQrCodeConfig.isRoomQrCodeEnabled
                QRCodeAcceptProximityInvitations       = $ConfigurationProfile.roomQrCodeConfig.isAcceptProximityInvitationEnabled
                #endregion

                #region Network Settings
                #All Devices
                DHCPEnabled                            = $ConfigurationProfile.dhcpEnabled
                LoggingEnabled                         = $ConfigurationProfile.loggingEnabled
                Hostname                               = $ConfigurationProfile.hostName
                DomainName                             = $ConfigurationProfile.domainName
                IPAddress                              = $ConfigurationProfile.ipAddress
                SubnetMask                             = $ConfigurationProfile.subnetMask
                DefaultGateway                         = $ConfigurationProfile.defaultGateway
                PrimaryDNS                             = $ConfigurationProfile.primaryDNS
                SecondaryDNS                           = $ConfigurationProfile.secondaryDNS
                DefaultAdminPassword                   = $ConfigurationProfile.deviceDefaultAdminPassword
                NetworkPCport                          = $ConfigurationProfile.networkPcPort
                #endregion

                #All Devices
                WhenCreated                            = (Get-Date "01/01/1970").AddSeconds($ConfigurationProfile.baseInfo.createdAt)
                CreatedByUser                          = $ConfigurationProfile.baseInfo.createdByUserName
                WhenChanged                            = (Get-Date "01/01/1970").AddSeconds($ConfigurationProfile.baseInfo.modifiedAt)
                ChangedByUser                          = $ConfigurationProfile.baseInfo.modifiedByUserName
            }
            $TDObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceConfigurationProfile')
            [void]$outTeamsDeviceConfiguration.Add($TDObj)
        }
        return $outTeamsDeviceConfiguration | Sort-Object DeviceType, DisplayName
    }
}

function Get-UcTeamsVersion {
    <#
        .SYNOPSIS
        Get Microsoft Teams Desktop Version

        .DESCRIPTION
        This function returns the installed Microsoft Teams desktop version for each user profile.

        .PARAMETER Path
        Specify the path with Teams Log Files

        .PARAMETER Computer
        Specify the remote computer

        .PARAMETER Credential
        Specify the credential to be used to connect to the remote computer

        .EXAMPLE
        PS> Get-UcTeamsVersion

        .EXAMPLE
        PS> Get-UcTeamsVersion -Path C:\Temp\

        .EXAMPLE
        PS> Get-UcTeamsVersion -Computer workstation124

        .EXAMPLE
        PS> $cred = Get-Credential
        PS> Get-UcTeamsVersion -Computer workstation124 -Credential $cred
    #>
    param(
        [string]$Path,
        [string]$Computer,
        [System.Management.Automation.PSCredential]$Credential,
        [switch]$SkipModuleCheck
    )
    
    $regexVersion = '("version":")([0-9.]*)'
    $regexRing = '("ring":")(\w*)'
    $regexEnv = '("environment":")(\w*)'
    $regexCloudEnv = '("cloudEnvironment":")(\w*)'
    
    $regexWindowsUser = '("upnWindowUserUpn":")([a-zA-Z0-9@._-]*)'
    $regexTeamsUserName = '("userName":")([a-zA-Z0-9@._-]*)'

    #2024-03-09: REGEX to get New Teams version from log file DesktopApp: Version: 23202.1500.2257.3700
    $regexNewVersion = '(DesktopApp: Version: )(\d{5}.\d{4}.\d{4}.\d{4})'
    
    #2025-01-31: Only need to check this once per PowerShell session
    $outTeamsVersion = [System.Collections.ArrayList]::new()
    if (!$SkipModuleCheck -and !$global:UCLobbyTeamsModuleCheck) {
        Test-UcPowerShellModule -ModuleName UcLobbyTeams | Out-Null
        $global:UCLobbyTeamsModuleCheck = $true
    }

    if ($Path) {
        if (Test-Path $Path -ErrorAction SilentlyContinue) {
            #region Teams Classic Path
            $TeamsSettingsFiles = Get-ChildItem -Path $Path -Include "settings.json" -Recurse
            foreach ($TeamsSettingsFile in $TeamsSettingsFiles) {
                $TeamsSettings = Get-Content -Path $TeamsSettingsFile.FullName
                $Version = ""
                $Ring = ""
                $Env = ""
                $CloudEnv = ""
                try {
                    $VersionTemp = [regex]::Match($TeamsSettings, $regexVersion).captures.groups
                    if ($VersionTemp.Count -ge 2) {
                        $Version = $VersionTemp[2].value
                    }
                    $RingTemp = [regex]::Match($TeamsSettings, $regexRing).captures.groups
                    if ($RingTemp.Count -ge 2) {
                        $Ring = $RingTemp[2].value
                    }
                    $EnvTemp = [regex]::Match($TeamsSettings, $regexEnv).captures.groups
                    if ($EnvTemp.Count -ge 2) {
                        $Env = $EnvTemp[2].value
                    }
                    $CloudEnvTemp = [regex]::Match($TeamsSettings, $regexCloudEnv).captures.groups
                    if ($CloudEnvTemp.Count -ge 2) {
                        $CloudEnv = $CloudEnvTemp[2].value
                    }
                }
                catch { }
                $TeamsDesktopSettingsFile = $TeamsSettingsFile.Directory.FullName + "\desktop-config.json"
                if (Test-Path $TeamsDesktopSettingsFile -ErrorAction SilentlyContinue) {
                    $TeamsDesktopSettings = Get-Content -Path $TeamsDesktopSettingsFile
                    $WindowsUser = ""
                    $TeamsUserName = ""
                    $RegexTemp = [regex]::Match($TeamsDesktopSettings, $regexWindowsUser).captures.groups
                    if ($RegexTemp.Count -ge 2) {
                        $WindowsUser = $RegexTemp[2].value
                    }
                    $RegexTemp = [regex]::Match($TeamsDesktopSettings, $regexTeamsUserName).captures.groups
                    if ($RegexTemp.Count -ge 2) {
                        $TeamsUserName = $RegexTemp[2].value
                    }
                }
                $TeamsVersion = New-Object -TypeName PSObject -Property @{
                    WindowsUser      = $WindowsUser
                    TeamsUser        = $TeamsUserName
                    Type             = "Teams Classic"
                    Version          = $Version
                    Ring             = $Ring
                    Environment      = $Env
                    CloudEnvironment = $CloudEnv
                    Path             = $TeamsSettingsFile.Directory.FullName
                }
                $TeamsVersion.PSObject.TypeNames.Insert(0, 'TeamsVersionFromPath')
                $outTeamsVersion.Add($TeamsVersion) | Out-Null
            }
            #endregion
            #region New Teams Path
            $TeamsSettingsFiles = Get-ChildItem -Path $Path -Include "tma_settings.json" -Recurse
            foreach ($TeamsSettingsFile in $TeamsSettingsFiles) {
                if (Test-Path $TeamsSettingsFile -ErrorAction SilentlyContinue) {
                    $NewTeamsSettings = Get-Content -Path $TeamsSettingsFile | ConvertFrom-Json
                    $tmpAccountID = $NewTeamsSettings.primary_user.accounts.account_id
                    try {
                        $Version = ""
                        $MostRecentTeamsLogFile = Get-ChildItem -Path $TeamsSettingsFile.Directory.FullName -Include "MSTeams_*.log" -Recurse | Sort-Object -Property CreationTime -Descending | Select-Object -First 1
                        $TeamLogContents = Get-Content $MostRecentTeamsLogFile
                        $RegexTemp = [regex]::Match($TeamLogContents, $regexNewVersion).captures.groups
                        if ($RegexTemp.Count -ge 2) {
                            $Version = $RegexTemp[2].value
                        }
                    }
                    catch {}

                    $TeamsVersion = New-Object -TypeName PSObject -Property @{
                        WindowsUser      = "NA"
                        TeamsUser        = $NewTeamsSettings.primary_user.accounts.account_upn
                        Type             = "New Teams"
                        Version          = $Version
                        Ring             = $NewTeamsSettings.tma_ecs_settings.$tmpAccountID.ring
                        Environment      = $NewTeamsSettings.tma_ecs_settings.$tmpAccountID.environment
                        CloudEnvironment = $NewTeamsSettings.primary_user.accounts.cloud
                        Path             = $TeamsSettingsFile.Directory.FullName
                    }
                    $TeamsVersion.PSObject.TypeNames.Insert(0, 'TeamsVersionFromPath')
                    [void]$outTeamsVersion.Add($TeamsVersion)
                }
            }
            #endregion
        }
        else {
            Write-Error -Message ("Invalid Path, please check if path: " + $path + " is correct and exists.")
        }
    }
    else {
        $currentDateFormat = [cultureinfo]::CurrentCulture.DateTimeFormat.ShortDatePattern
        if ($Computer) {
            $RemotePath = "\\" + $Computer + "\C$\Users"
            $ComputerName = $Computer
            if ($Credential) {
                if ($Computer.IndexOf('.') -gt 0) {
                    $PSDriveName = $Computer.Substring(0, $Computer.IndexOf('.')) + "_TmpTeamsVersion"
                }
                else {
                    $PSDriveName = $Computer + "_TmpTeamsVersion"
                }
                New-PSDrive -Root $RemotePath -Name $PSDriveName -PSProvider FileSystem -Credential $Credential | Out-Null
            }

            if (Test-Path -Path $RemotePath) {
                $Profiles = Get-ChildItem -Path $RemotePath -ErrorAction SilentlyContinue
            }
            else {
                Write-Error -Message ("Error: Cannot get users on " + $computer + ", please check if name is correct and if the current user has permissions.")
            }
        }
        else {
            $ComputerName = $Env:COMPUTERNAME
            $Profiles = Get-childItem 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' | ForEach-Object { Get-ItemProperty $_.pspath } | Where-Object { $_.fullprofile -eq 1 }
            $newTeamsFound = $false
        }
       
        foreach ($UserProfile in $Profiles) {
            if ($Computer) {
                $ProfilePath = $UserProfile.FullName
                $ProfileName = $UserProfile.Name
            }
            else {
                $ProfilePath = $UserProfile.ProfileImagePath
                #2023-10-13: Added exception handeling, only known case is when a windows profile was created when the machine was joined to a previous domain.
                try {
                    $ProfileName = (New-Object System.Security.Principal.SecurityIdentifier($UserProfile.PSChildName)).Translate( [System.Security.Principal.NTAccount]).Value
                }
                catch {
                    $ProfileName = "Unknown Windows User"
                }
            }
            #region classic teams
            #2024-10-30: We will only add an entry if the executable exists.
            $TeamsApp = $ProfilePath + "\AppData\Local\Microsoft\Teams\current\Teams.exe"
            if (Test-Path -Path $TeamsApp) {
                $TeamsSettingPath = $ProfilePath + "\AppData\Roaming\Microsoft\Teams\settings.json"
                if (Test-Path $TeamsSettingPath -ErrorAction SilentlyContinue) {
                    $TeamsSettings = Get-Content -Path $TeamsSettingPath
                    $Version = ""
                    $Ring = ""
                    $Env = ""
                    $CloudEnv = ""
                    try {
                        $VersionTemp = [regex]::Match($TeamsSettings, $regexVersion).captures.groups
                        if ($VersionTemp.Count -ge 2) {
                            $Version = $VersionTemp[2].value
                        }
                        $RingTemp = [regex]::Match($TeamsSettings, $regexRing).captures.groups
                        if ($RingTemp.Count -ge 2) {
                            $Ring = $RingTemp[2].value
                        }
                        $EnvTemp = [regex]::Match($TeamsSettings, $regexEnv).captures.groups
                        if ($EnvTemp.Count -ge 2) {
                            $Env = $EnvTemp[2].value
                        }
                        $CloudEnvTemp = [regex]::Match($TeamsSettings, $regexCloudEnv).captures.groups
                        if ($CloudEnvTemp.Count -ge 2) {
                            $CloudEnv = $CloudEnvTemp[2].value
                        }
                    }
                    catch { }
                    $TeamsInstallTimePath = $ProfilePath + "\AppData\Roaming\Microsoft\Teams\installTime.txt"
                    #2024-02-28: In some cases the install file can be missing.
                    $tmpInstallDate = ""
                    if (Test-Path $TeamsInstallTimePath -ErrorAction SilentlyContinue) {
                        $InstallDateStr = Get-Content ($ProfilePath + "\AppData\Roaming\Microsoft\Teams\installTime.txt")
                        $tmpInstallDate = [Datetime]::ParseExact($InstallDateStr, 'M/d/yyyy', $null) | Get-Date -Format $currentDateFormat
                    }
                    $TeamsVersion = New-Object -TypeName PSObject -Property @{
                        Computer         = $ComputerName
                        Profile          = $ProfileName
                        ProfilePath      = $ProfilePath
                        Type             = "Teams Classic"
                        Version          = $Version
                        Ring             = $Ring
                        Environment      = $Env
                        CloudEnvironment = $CloudEnv
                        Arch             = Get-UcArch $TeamsApp
                        InstallDate      = $tmpInstallDate
                    }
                    $TeamsVersion.PSObject.TypeNames.Insert(0, 'TeamsVersion')
                    [void]$outTeamsVersion.Add($TeamsVersion)
                }
            }
            #endregion

            #region New Teams
            #2024-10-28: Running this in Windows 10 with PowerShell 7 an exception could be raised while importing the Appx PowerShell module. Thank you Steve Chupack for reporting this issue.
            $newTeamsLocation = ""
            if ($Computer) {
                $newTeamsLocation = Get-ChildItem -Path ( $RemotePath + "\..\Program Files\Windowsapps" ) -Filter "ms-teams.exe" -Recurse -Depth 1 | Sort-Object -Property CreationTime -Descending | Select-Object -First 1
            }
            else {
                #If running as Administrator then we can search instead, this is to prevent using Get-AppPackage MSTeams -AllUser which also requires Administrator
                if (Test-UcElevatedPrivileges) {
                    $newTeamsLocation = Get-ChildItem -Path "C:\Program Files\Windowsapps" -Filter "ms-teams.exe" -Recurse -Depth 1 | Sort-Object -Property CreationTime -Descending | Select-Object -First 1
                }
                else {
                    try {
                        #Checking if the module is already loaded
                        if (!(Get-Module Appx)) {
                            Import-Module Appx
                        }
                    }
                    catch [System.PlatformNotSupportedException] {
                        Import-Module Appx -UseWindowsPowerShell
                    }
                    $TeamsAppPackage = Get-AppPackage MSTeams
                    if ($TeamsAppPackage) {
                        $newTeamsInstallPath = $TeamsAppPackage.InstallLocation + ".\ms-teams.exe"
                        $newTeamsLocation = Get-ItemProperty -Path ($newTeamsInstallPath)
                    }
                }
            }
            if ($newTeamsLocation) {
                if (Test-Path -Path $newTeamsLocation.FullName -ErrorAction SilentlyContinue) {
                    $tmpRing = ""
                    $tmpEnvironment = ""
                    $tmpCloudEnvironment = ""
                    $NewTeamsSettingPath = $ProfilePath + "\AppData\Local\Publishers\8wekyb3d8bbwe\TeamsSharedConfig\tma_settings.json"
                    if (Test-Path $NewTeamsSettingPath -ErrorAction SilentlyContinue) {
                        try {
                            $NewTeamsSettings = Get-Content -Path $NewTeamsSettingPath | ConvertFrom-Json
                            $tmpAccountID = $NewTeamsSettings.primary_user.accounts.account_id
                            $tmpRing = $NewTeamsSettings.tma_ecs_settings.$tmpAccountID.ring
                            $tmpEnvironment = $NewTeamsSettings.tma_ecs_settings.$tmpAccountID.environment
                            $tmpCloudEnvironment = $NewTeamsSettings.primary_user.accounts.cloud
                        }
                        catch {}
                    }
                    $TeamsVersion = New-Object -TypeName PSObject -Property @{
                        Computer         = $ComputerName
                        Profile          = $ProfileName
                        ProfilePath      = $ProfilePath
                        Type             = "New Teams"
                        Version          = $newTeamsLocation.VersionInfo.ProductVersion
                        Ring             = $tmpRing
                        Environment      = $tmpEnvironment
                        CloudEnvironment = $tmpCloudEnvironment
                        Arch             = Get-UcArch $newTeamsLocation.FullName
                        InstallDate      = $newTeamsLocation.CreationTime | Get-Date -Format $currentDateFormat
                    }
                    $TeamsVersion.PSObject.TypeNames.Insert(0, 'TeamsVersion')
                    [void]$outTeamsVersion.Add($TeamsVersion)
                    $newTeamsFound = $true
                }
            }
            #endregion
        }
        if ($Credential -and $PSDriveName) {
            try {
                Remove-PSDrive -Name $PSDriveName -ErrorAction SilentlyContinue
            }
            catch {}
        }
    }
    if (!(Test-UcElevatedPrivileges) -and !($Computer) -and !($newTeamsFound)) {
        Write-Warning "No New Teams versions found, please try again with elevated privileges (Run as Administrator)"
    }
    return $outTeamsVersion
}

function Get-UcTeamsVersionBatch {
    <#
        .SYNOPSIS
        Get Microsoft Teams Desktop Version from all computers in a csv file.

        .DESCRIPTION
        This function returns the installed Microsoft Teams desktop version for each user profile.

        .PARAMETER InputCSV
        CSV with the list of computers that we want to get the Teams Version

        .PARAMETER OutputPath
        Specify the output path

        .PARAMETER ExportCSV
        Export the output to a CSV file

        .PARAMETER Credential
        Specify the credential to be used to connect to the remote computers

        .EXAMPLE
        PS> Get-UcTeamsVersionBatch

        .EXAMPLE
        PS> Get-UcTeamsVersionBatch -InputCSV C:\Temp\ComputerList.csv -Credential $cred

        .EXAMPLE
        PS> Get-UcTeamsVersionBatch -InputCSV C:\Temp\ComputerList.csv -Credential $cred -ExportCSV
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputCSV,
        [string]$OutputPath,
        [switch]$ExportCSV,
        [System.Management.Automation.PSCredential]$Credential
    )

    #2025-01-31: Only need to check this once per PowerShell session
    if (!($global:UCLobbyTeamsModuleCheck)) {
        Test-UcPowerShellModule -ModuleName UcLobbyTeams | Out-Null
        $global:UCLobbyTeamsModuleCheck = $true
    }
    if (Test-Path $InputCSV) {
        try {
            $Computers = Import-Csv -Path $InputCSV
        }
        catch {
            Write-Host ("Invalid CSV input file: " + $InputCSV) -ForegroundColor Red
            return
        }
        $outTeamsVersion = [System.Collections.ArrayList]::new()
        #Verify if the Output Path exists
        if ($OutputPath) {
            if (!(Test-Path $OutputPath -PathType Container)) {
                Write-Host ("Error: Invalid folder: " + $OutputPath) -ForegroundColor Red
                return
            } 
        }
        else {                
            $OutputPath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads")
        }

        $c = 0
        $compCount = $Computers.count
        
        foreach ($computer in $Computers) {
            $c++
            Write-Progress -Activity ("Getting Teams Version from: " + $computer.Computer)  -Status "Computer $c of $compCount "
            $tmpTV = Get-UcTeamsVersion -Computer $computer.Computer -Credential $cred -SkipModuleCheck
            $outTeamsVersion.Add($tmpTV) | Out-Null
        }
        if ($ExportCSV) {
            $tmpFileName = "MSTeamsVersion_" + ( get-date ).ToString('yyyyMMdd-HHmmss') + ".csv"
            $OutputFullPath = [System.IO.Path]::Combine($OutputPath, $tmpFileName)
            $outTeamsVersion | Sort-Object Computer, Profile | Select-Object Computer, Profile, ProfilePath, Arch, Version, Environment, Ring, InstallDate | Export-Csv -path $OutputFullPath -NoTypeInformation
            Write-Host ("Results available in: " + $OutputFullPath) -ForegroundColor Cyan
        }
        else {
            return $outTeamsVersion 
        }
    }
    else {
        Write-Host ("Error: File not found " + $InputCSV) -ForegroundColor Red
    }
}

function Get-UcTeamsWithSingleOwner {
    <#
        .SYNOPSIS
        Get Teams that have a single owner

        .DESCRIPTION
        This function returns a list of Teams that only have a single owner.

        Requirements:   Microsoft Teams PowerShell Module (Install-Module MicrosoftTeams)

        .EXAMPLE
        PS> Get-UcTeamsWithSingleOwner
    #>
    Get-UcTeamUsersEmail -Role Owner -Confirm:$false | Group-Object -Property TeamDisplayName | Where-Object { $_.Count -lt 2 } | Select-Object -ExpandProperty Group
}

function Get-UcTeamUsersEmail {
    <#
        .SYNOPSIS
        Get Users Email Address that are in a Team

        .DESCRIPTION
        This function returns a list of users email address that are part of a Team.

        Requirements:   Microsoft Teams PowerShell Module (Install-Module MicrosoftTeams)

        .PARAMETER TeamName
        Specifies Team Name.

        .PARAMETER Role
        Specifies which roles to filter (Owner, User, Guest)

        .EXAMPLE
        PS> Get-UcTeamUsersEmail

        .EXAMPLE
        PS> Get-UcTeamUsersEmail -TeamName "Marketing"

        .EXAMPLE
        PS> Get-UcTeamUsersEmail -Role "Guest"

        .EXAMPLE
        PS> Get-UcTeamUsersEmail -TeamName "Marketing" -Role "Guest"
    #>
    [cmdletbinding(SupportsShouldProcess)]
    param(
        [string]$TeamName,
        [ValidateSet("Owner", "User", "Guest")] 
        [string]$Role
    )
    
    #region 2025-03-31: Check if connected with Teams PowerShell module. 
    if (!(Test-UcServiceConnection -Type TeamsPowerShell)) {
        return
    }
    #2025-01-31: Only need to check this once per PowerShell session
    if (!($global:UCLobbyTeamsModuleCheck)) {
        Test-UcPowerShellModule -ModuleName UcLobbyTeams | Out-Null
        $global:UCLobbyTeamsModuleCheck = $true
    }
    #endregion

    $output = [System.Collections.ArrayList]::new()
    if ($TeamName) {
        $Teams = Get-Team -DisplayName $TeamName
    }
    else {
        if ($ConfirmPreference) {
            $title = 'Confirm'
            $question = 'Are you sure that you want to list all Teams?'
            $choices = '&Yes', '&No'
            $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
        }
        else {
            $decision = 0
        }
        if ($decision -eq 0) {
            $Teams = Get-Team
        }
        else {
            return
        }
    }
    foreach ($Team in $Teams) { 
        if ($Role) {
            $TeamMembers = Get-TeamUser -GroupId $Team.GroupID -Role $Role
        }
        else {
            $TeamMembers = Get-TeamUser -GroupId $Team.GroupID 
        }
        foreach ($TeamMember in $TeamMembers) {
            $Email = ( Get-csOnlineUser $TeamMember.User | Select-Object @{Name = 'PrimarySMTPAddress'; Expression = { $_.ProxyAddresses -cmatch '^SMTP:' -creplace 'SMTP:' } }).PrimarySMTPAddress
            $Member = [PSCustomObject][Ordered]@{
                TeamGroupID     = $Team.GroupID
                TeamDisplayName = $Team.DisplayName
                TeamVisibility  = $Team.Visibility
                UPN             = $TeamMember.User
                Role            = $TeamMember.Role
                Email           = $Email
            }
            $Member.PSObject.TypeNames.Insert(0, 'TeamUsersEmail')
            [void]$output.Add($Member) 
        }
    }
    return $output
}

function Set-UcTeamsDeviceConfigurationProfile {
    <#
        .SYNOPSIS
        Allow assign a Teams Device Configuration Profile to one or more Teams Devices

        .DESCRIPTION
        This function will use TAC API to assign a Configuration Profile by sending the Config Update command to the specified device(s).

        Requirements:   EntraAuth PowerShell Module (Install-Module EntraAuth)

        .PARAMETER TACDeviceID
        Teams Device ID from Teams Admin Center.

        .PARAMETER ConfigID
        Teams Device Configuration Profile ID.

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
    if (Test-UcServiceConnection -Type TeamsDeviceTAC) {
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

function Test-UcPowerShellModule {
    <#
        .SYNOPSIS
        Test if PowerShell module is installed and updated

        .DESCRIPTION
        This function returns FALSE if PowerShell module is not installed.

        .PARAMETER ModuleName
        Specifies PowerShell module name

        .EXAMPLE
        PS> Test-UcPowerShellModule -ModuleName UCLobbyTeams
    #>
    param(
        [Parameter(Mandatory = $true)]    
        [string]$ModuleName
    )
    
    try {
        #Region 2025-07-23: We can use the current module name, this will make the code simpler in the other functions.
        $ModuleName = $MyInvocation.MyCommand.Module.Name
        if (!($ModuleName)) {
            Write-Warning "Please specify a module name using the ModuleName parameter."
            return
        }
        $ModuleNameCheck = Get-Variable -Scope Global -Name ($ModuleName + "ModuleCheck") -ErrorAction SilentlyContinue
        if ($ModuleNameCheck.Value) {
            return $true
        }
        if ($ModuleNameCheck) {
            Set-Variable -Scope Global -Name ($ModuleName + "ModuleCheck") -Value $true
        }
        else {
            New-Variable -Scope Global -Name ($ModuleName + "ModuleCheck") -Value $true
        }
        #endRegion
         
        #Get all installed versions
        $installedVersions = (Get-Module $ModuleName -ListAvailable | Sort-Object Version -Descending).Version

        #Get the lastest version available
        $availableVersion = (Find-Module -Name $ModuleName -Repository PSGallery -ErrorAction SilentlyContinue).Version

        if (!($installedVersions)) {
            if ($availableVersion ) {
                #Module not installed and there is an available version to install.
                Write-Warning ("The PowerShell Module $ModuleName is not installed, please install the latest available version ($availableVersion) with:" + [Environment]::NewLine + "Install-Module $ModuleName")
            }
            else {
                #Wrong name or not found in the registered PS Repository.
                Write-Warning ("The PowerShell Module $ModuleName not found in the registered PS Repository, please check the module name and try again.")
            }
            return $false
        }

        #Get the current loaded version
        $tmpCurrentVersion = (Get-Module $ModuleName | Sort-Object Version -Descending)
        if ($tmpCurrentVersion) {
            $currentVersion = $tmpCurrentVersion[0].Version.ToString()
        }

        if (!($currentVersion)) {
            #Module is installed but not imported, in this case we check if there is a newer version available.
            if ($availableVersion -in $installedVersions) {
                Write-Warning ("The lastest available version of $ModuleName module is installed, however the module is not imported." + [Environment]::NewLine + "Please make sure you import it with:" + [Environment]::NewLine + "Import-Module $ModuleName -RequiredVersion $availableVersion")
                return $false
            }
            else {
                Write-Warning ("There is a new version available $availableVersion, the lastest installed version is " + $installedVersions[0] + "." + [Environment]::NewLine + "Please update the module with:" + [Environment]::NewLine + "Update-Module $ModuleName")
            }
        }

        if ($currentVersion -ne $availableVersion ) {
            if ($availableVersion -in $installedVersions) {
                Write-Warning ("The lastest available version of $ModuleName module is installed, however version $currentVersion is imported." + [Environment]::NewLine + "Please make sure you import it with:" + [Environment]::NewLine + "Import-Module $ModuleName -RequiredVersion $availableVersion")
            }
            else {
                Write-Warning ("There is a new version available $availableVersion, current version $currentVersion." + [Environment]::NewLine + "Please update the module with:" + [Environment]::NewLine + "Update-Module $ModuleName")
            }
        }
        return $true
    }
    catch {
    }
    return $false
}

function Test-UcTeamsDevicesCompliancePolicy {
    <#
        .SYNOPSIS
        Validate which Intune Compliance policies are supported by Microsoft Teams Android Devices

        .DESCRIPTION
        This function will validate each setting in the Intune Compliance Policy to make sure they are in line with the supported settings:

            https://docs.microsoft.com/en-us/microsoftteams/rooms/supported-ca-and-compliance-policies?tabs=phones#supported-device-compliance-policies

        Contributors: Traci Herr, David Paulino and Gonçalo Sepúlveda
        
        Requirements:   EntraAuth PowerShell Module (Install-Module EntraAuth)
                        or
                        Microsoft Graph Authentication PowerShell Module (Install-Module Microsoft.Graph.Authentication)

                        Microsoft Graph Scopes:
                            "DeviceManagementConfiguration.Read.All", "Directory.Read.All"

        .PARAMETER PolicyID
        Specifies a Policy ID that will be checked if is supported by Microsoft Teams Devices.

        .PARAMETER PolicyName
        Specifies a Policy Name that will be checked if is supported by Microsoft Teams Devices
                
        .PARAMETER UserUPN
        Specifies a UserUPN that we want to check for applied compliance policies.
        
        .PARAMETER DeviceID
        Specifies DeviceID that we want to check for applied compliance policies.
        
        .PARAMETER Detailed
        Displays test results for unsupported settings in each Intune Compliance Policy

        .PARAMETER All
        Will check all Intune Compliance policies independently if they are assigned to a User(s)/Group(s).

        .PARAMETER IncludeSupported
        Displays results for all settings in each Intune Compliance Policy.

        .PARAMETER ExportCSV
        When present will export the detailed results to a CSV file. By defautl will save the file under the current user downloads, unless we specify the OutputPath.

        .PARAMETER OutputPath
        Allows to specify the path where we want to save the results.

        .EXAMPLE 
        PS> Test-UcTeamsDevicesCompliancePolicy

        .EXAMPLE 
        PS> Test-UcTeamsDevicesCompliancePolicy -Detailed
    #>
    param(
        [string]$PolicyID,
        [string]$PolicyName,
        [string]$UserUPN,
        [string]$DeviceID,
        [switch]$Detailed,
        [switch]$All,
        [switch]$IncludeSupported,
        [switch]$ExportCSV,
        [string]$OutputPath
    )

    $CompliancePolicies = $null
    $totalCompliancePolicies = 0
    $skippedCompliancePolicies = 0
    $output = [System.Collections.ArrayList]::new()
    $outputSum = [System.Collections.ArrayList]::new()
    $Groups = New-Object 'System.Collections.Generic.Dictionary[string, string]'

    $GraphPathUsers = "/users"
    $GraphPathGroups = "/groups"
    $GraphPathCompliancePolicies = "/deviceManagement/deviceCompliancePolicies/?`$expand=assignments"
    $GraphPathDevices = "/devices"

    $SupportedAndroidCompliancePolicies = "#microsoft.graph.androidCompliancePolicy", "#microsoft.graph.androidDeviceOwnerCompliancePolicy", "#microsoft.graph.aospDeviceOwnerCompliancePolicy"
    $SupportedWindowsCompliancePolicies = "#microsoft.graph.windows10CompliancePolicy"

    $URLSupportedCompliancePoliciesAndroid = "https://aka.ms/TeamsDevicePolicies?tabs=phones-da&#supported-device-compliance-policies"
    $URLSupportedCompliancePoliciesAOSP = "https://aka.ms/TeamsDevicePolicies?tabs=phones&#supported-device-compliance-policies"
    $URLSupportedCompliancePoliciesWindows = "https://aka.ms/TeamsDevicePolicies?tabs=mtr-w&#supported-device-compliance-policies"

    #region Graph Connection, Scope validation and module version
    if (!(Test-UcServiceConnection -Type MSGraph -Scopes "DeviceManagementConfiguration.Read.All", "Directory.Read.All")) {
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
        $outFileName = "TeamsDevices_CompliancePolicy_Report_" + ( get-date ).ToString('yyyyMMdd-HHmmss') + ".csv"
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
        Write-Progress -Activity "Test-UcTeamsDeviceCompliancePolicy" -Status "Getting Compliance Policies"
        $CompliancePolicies = Invoke-UcGraphRequest -Path $GraphPathCompliancePolicies -Beta
        if ($UserUPN) {
            try {
                $UserGroups = (Invoke-UcGraphRequest -Path ($GraphPathUsers + "/" + $userUPN + "/transitiveMemberOf?`$select=id")).id
            }
            catch [System.Net.Http.HttpRequestException] {
                if ($PSItem.Exception.Response.StatusCode -eq "NotFound") {
                    Write-warning -Message ("User Not Found: " + $UserUPN)
                }
                return
            }
            #We also need to take in consideration devices that are registered to this user
            $DeviceGroups = [System.Collections.ArrayList]::new()
            $userDevices = Invoke-UcGraphRequest -Path ($GraphPathUsers + "/" + $userUPN + "/registeredDevices?`$select=deviceId,displayName") 
            foreach ($userDevice in $userDevices) {
                $tmpGroups = (Invoke-UcGraphRequest -Path ($GraphPathDevices + "(deviceId='{" + $userDevice.deviceID + "}')/transitiveMemberOf?`$select=id")).id
                foreach ($tmpGroup in $tmpGroups) {
                    $tmpDG = New-Object -TypeName PSObject -Property @{
                        GroupId           = $tmpGroup
                        DeviceId          = $userDevice.deviceID
                        DeviceDisplayName = $userDevice.displayName
                    }
                    [void]$DeviceGroups.Add($tmpDG)
                }
            }
        }
        elseif ($DeviceID) {
            try {
                $DeviceGroups = (Invoke-UcGraphRequest -Path ($GraphPathDevices + "(deviceId='{" + $DeviceID + "}')/transitiveMemberOf?`$select=id")).id
            }
            catch [System.Net.Http.HttpRequestException] {
                if ($PSItem.Exception.Response.StatusCode -eq "BadRequest") {
                    Write-warning -Message ("Device ID Not Found: " + $DeviceID)
                }
                return
            }
        }
    }
    catch [System.Net.Http.HttpRequestException] {
        if ($PSItem.Exception.Response.StatusCode -eq "Unauthorized") {
            Write-Error "Access Denied, please make sure the user connecing to MS Graph is part of one of the following Global Reader/Intune Service Administrator/Global Administrator roles"
        }
        else {
            Write-Error $PSItem.Exception.Message
        }
    }
    catch {
        Write-Error $PSItem.Exception.Message
    }

    $p = 0
    foreach ($CompliancePolicy in $CompliancePolicies) {
        $p++
        Write-Progress -Activity "Test-UcTeamsDeviceCompliancePolicy" -Status ("Checking policy " + $CompliancePolicy.displayName + " - $p of " + $CompliancePolicies.Count)
        if ((($PolicyID -eq $CompliancePolicy.id) `
                    -or ($PolicyName -eq $CompliancePolicy.displayName) `
                    -or (!$PolicyID -and !$PolicyName)) `
                -and (($CompliancePolicy."@odata.type" -in $SupportedAndroidCompliancePolicies) `
                    -or ($CompliancePolicy."@odata.type" -in $SupportedWindowsCompliancePolicies))) {
            
            $AssignedToGroup = [System.Collections.ArrayList]::new()
            $ExcludedFromGroup = [System.Collections.ArrayList]::new()
            $outAssignedToGroup = ""
            $outExcludedFromGroup = ""
            #We wont need to check the settings if the policy is not assigned to a user
            if ($UserUPN -or $DeviceID) {
                $userOrDeviceIncluded = $false
            }
            else {
                $userOrDeviceIncluded = $true
            }

            #Checking Compliance Policy assigments since we can skip non assigned policies.
            foreach ($CompliancePolicyAssignment in $CompliancePolicy.assignments) {
                $GroupDisplayName = $CompliancePolicyAssignment.target.Groupid
                if ($Groups.ContainsKey($CompliancePolicyAssignment.target.Groupid)) {
                    $GroupDisplayName = $Groups.Item($CompliancePolicyAssignment.target.Groupid)
                }
                else {
                    try {
                        $GroupInfo = Invoke-UcGraphRequest -Path ($GraphPathGroups + "/" + $CompliancePolicyAssignment.target.Groupid + "/?`$select=id,displayname")
                        $Groups.Add($GroupInfo.id, $GroupInfo.displayname)
                        $GroupDisplayName = $GroupInfo.displayname
                    }
                    catch {
                    }
                }
                $GroupEntry = [PSCustomObject][Ordered]@{
                    GroupID          = $CompliancePolicyAssignment.target.Groupid
                    GroupDisplayName = $GroupDisplayName
                }
                switch ($CompliancePolicyAssignment.target."@odata.type") {
                    #Policy assigned to all users
                    "#microsoft.graph.allLicensedUsersAssignmentTarget" {
                        $GroupEntry = [PSCustomObject][Ordered]@{
                            GroupID          = "allLicensedUsersAssignment"
                            GroupDisplayName = "All Users"
                        }
                        [void]$AssignedToGroup.Add($GroupEntry)
                        $userOrDeviceIncluded = $true
                    }
                    #Policy assigned to all devices
                    "#microsoft.graph.allDevicesAssignmentTarget" {
                        $GroupEntry = New-Object -TypeName PSObject -Property @{
                            GroupID          = "allDevicesAssignmentTarget"
                            GroupDisplayName = "All Devices"
                        }
                        [void]$AssignedToGroup.Add($GroupEntry) 
                        $userOrDeviceIncluded = $true
                    }
                    #Group that this policy is assigned
                    "#microsoft.graph.groupAssignmentTarget" {
                        [void]$AssignedToGroup.Add($GroupEntry) 
                        if (($UserUPN -or $DeviceID) -and (($CompliancePolicyAssignment.target.Groupid -in $UserGroups) -or ($CompliancePolicyAssignment.target.Groupid -in $DeviceGroups))) {
                            $userOrDeviceIncluded = $true
                        }
                    }
                    #Group that this policy is excluded
                    "#microsoft.graph.exclusionGroupAssignmentTarget" {
                        [void]$ExcludedFromGroup.Add($GroupEntry)
                        #If user is excluded then we dont need to check the policy
                        if ($UserUPN -and ($CompliancePolicyAssignment.target.Groupid -in $UserGroups)) {
                            Write-Warning ("Skiping compliance policy " + $CompliancePolicy.displayName + ", since user " + $UserUPN + " is part of an Excluded Group: " + $GroupEntry.GroupDisplayName)
                            $userOrDeviceExcluded = $true
                        }
                        elseif ($DeviceID -and ($CompliancePolicyAssignment.target.Groupid -in $DeviceGroups)) {
                            Write-Warning ("Skiping compliance policy " + $CompliancePolicy.displayName + ", since device " + $DeviceID + " is part of an Excluded Group: " + $GroupEntry.GroupDisplayName)
                            $userOrDeviceExcluded = $true
                        }
                        elseif ($UserUPN -and (($CompliancePolicyAssignment.target.Groupid -in $DeviceGroups.GroupId))) {
                            #In case a device is excluded we will check the policy but output a message
                            $tmpDev = ($DeviceGroups | Where-Object -Property GroupId -eq -Value $CompliancePolicyAssignment.target.Groupid)
                            Write-Warning ("Compliance policy " + $CompliancePolicy.displayName + " will not be applied to device " + $tmpDev.DeviceDisplayName + " (" + $tmpDev.DeviceID + "), since this device is part of an Excluded Group: " + $GroupEntry.GroupDisplayName)
                        }
                    }
                }
            }
                
            if ((($AssignedToGroup.count -gt 0) -and !$userOrDeviceExcluded -and $userOrDeviceIncluded) -or $all) {
                $totalCompliancePolicies++ 
                $PolicyErrors = 0
                $PolicyWarnings = 0
                #Define the Compliance Policy type
                switch ($CompliancePolicy."@odata.type") {
                    "#microsoft.graph.androidCompliancePolicy" { $CPType = "Android Device"; $URLSupportedCompliancePolicies = $URLSupportedCompliancePoliciesAndroid }
                    "#microsoft.graph.androidDeviceOwnerCompliancePolicy" { $CPType = "Android Enterprise"; $URLSupportedCompliancePolicies = $URLSupportedCompliancePoliciesAndroid }
                    "#microsoft.graph.aospDeviceOwnerCompliancePolicy" { $CPType = "Android (AOSP)"; $URLSupportedCompliancePolicies = $URLSupportedCompliancePoliciesAOSP }
                    "#microsoft.graph.windows10CompliancePolicy" { $CPType = "Windows 10 or later"; $URLSupportedCompliancePolicies = $URLSupportedCompliancePoliciesWindows }
                    Default { $CPType = $CompliancePolicy."@odata.type".split('.')[2] }
                }

                #If only assigned/excluded from a group we will show the group display name, otherwise the number of groups assigned/excluded.
                if ($AssignedToGroup.count -eq 1) {
                    $outAssignedToGroup = $AssignedToGroup.GroupDisplayName
                }
                elseif ($AssignedToGroup.count -eq 0) {
                    $outAssignedToGroup = "None"
                }
                else {
                    $outAssignedToGroup = "" + $AssignedToGroup.count + " groups"
                }
                if ($ExcludedFromGroup.count -eq 1) {
                    $outExcludedFromGroup = $ExcludedFromGroup.GroupDisplayName
                }
                elseif ($ExcludedFromGroup.count -eq 0) {
                    $outExcludedFromGroup = "None"
                }
                else {
                    $outExcludedFromGroup = "" + $ExcludedFromGroup.count + " groups"
                }

                #region Common settings between Android (ADA and AOSP) and Windows 
                #region 9: Device Properties > Operation System Version
                $ID = 9.1
                $Setting = "osMinimumVersion"
                $SettingDescription = "Device Properties > Operation System Version > Minimum OS version"
                $SettingValue = "Not Configured"
                $Comment = ""
                $Status = "Supported"
                if (!([string]::IsNullOrEmpty($CompliancePolicy.osMinimumVersion))) {
                    if ($CompliancePolicy."@odata.type" -in $SupportedWindowsCompliancePolicies) {
                        $Status = "Unsupported"
                        $Comment = "Teams Rooms automatically updates to newer versions of Windows and setting values here could prevent successful sign-in after an OS update."
                        $PolicyErrors++
                    }
                    else {
                        $Status = "Warning"
                        $Comment = "This setting can cause sign in issues."
                        $PolicyWarnings++
                    }
                    $SettingValue = $CompliancePolicy.osMinimumVersion
                }
                $SettingPSObj = [PSCustomObject][Ordered]@{
                    ID                    = $ID
                    PolicyName            = $CompliancePolicy.displayName
                    PolicyID              = $CompliancePolicy.id
                    PolicyType            = $CPType
                    AssignedToGroup       = $outAssignedToGroup
                    AssignedToGroupList   = $AssignedToGroup
                    ExcludedFromGroup     = $outExcludedFromGroup 
                    ExcludedFromGroupList = $ExcludedFromGroup
                    TeamsDevicesStatus    = $Status 
                    Setting               = $Setting
                    Value                 = $SettingValue
                    SettingDescription    = $SettingDescription
                    Comment               = $Comment
                }
                [void]$output.Add($SettingPSObj)
            
                $ID = 9.2
                $Setting = "osMaximumVersion"
                $SettingDescription = "Device Properties > Operation System Version > Maximum OS version"
                $SettingValue = "Not Configured"
                $Comment = ""
                $Status = "Supported"
                if (!([string]::IsNullOrEmpty($CompliancePolicy.osMaximumVersion))) {
                    if ($CompliancePolicy."@odata.type" -in $SupportedWindowsCompliancePolicies) {
                        $Status = "Unsupported"
                        $Comment = "Teams Rooms automatically updates to newer versions of Windows and setting values here could prevent successful sign-in after an OS update."
                        $PolicyErrors++
                    }
                    else {
                        $Status = "Warning"
                        $Comment = "This setting can cause sign in issues."
                        $PolicyWarnings++
                    }
                    $SettingValue = $CompliancePolicy.osMaximumVersion
                }
                $SettingPSObj = [PSCustomObject][Ordered]@{
                    ID                    = $ID
                    PolicyName            = $CompliancePolicy.displayName
                    PolicyID              = $CompliancePolicy.id
                    PolicyType            = $CPType
                    AssignedToGroup       = $outAssignedToGroup
                    AssignedToGroupList   = $AssignedToGroup
                    ExcludedFromGroup     = $outExcludedFromGroup 
                    ExcludedFromGroupList = $ExcludedFromGroup
                    TeamsDevicesStatus    = $Status 
                    Setting               = $Setting
                    Value                 = $SettingValue
                    SettingDescription    = $SettingDescription
                    Comment               = $Comment
                }
                [void]$output.Add($SettingPSObj)
                #endregion

                #region 17: System Security > All Android devices > Require a password to unlock mobile devices
                $ID = 17
                $Setting = "passwordRequired"
                $SettingDescription = "System Security > All Android devices > Require a password to unlock mobile devices"
                $SettingValue = "Not Configured"
                $Status = "Supported"
                $Comment = ""
                if ($CompliancePolicy.passwordRequired) {
                    $Status = "Unsupported"
                    $SettingValue = "Require"
                    $Comment = $URLSupportedCompliancePolicies
                    $PolicyErrors++
                }
                $SettingPSObj = [PSCustomObject][Ordered]@{
                    ID                    = $ID
                    PolicyName            = $CompliancePolicy.displayName
                    PolicyID              = $CompliancePolicy.id
                    PolicyType            = $CPType
                    AssignedToGroup       = $outAssignedToGroup
                    AssignedToGroupList   = $AssignedToGroup
                    ExcludedFromGroup     = $outExcludedFromGroup 
                    ExcludedFromGroupList = $ExcludedFromGroup
                    TeamsDevicesStatus    = $Status 
                    Setting               = $Setting
                    Value                 = $SettingValue
                    SettingDescription    = $SettingDescription
                    Comment               = $Comment
                }
                [void]$output.Add($SettingPSObj)
                #endregion
                #endregion

                #2024-05-08 - We need to limit the settings since not all are available in AOSP compliance policies.
                if ($CompliancePolicy."@odata.type" -in $SupportedAndroidCompliancePolicies -and $CompliancePolicy."@odata.type" -ne "#microsoft.graph.aospDeviceOwnerCompliancePolicy") {

                    #region 1: Microsoft Defender for Endpoint > Require the device to be at or under the machine risk score
                    $ID = 1
                    $Setting = "deviceThreatProtectionEnabled"
                    $SettingDescription = "Microsoft Defender for Endpoint > Require the device to be at or under the machine risk score"
                    $SettingValue = "Not Configured"
                    $Comment = ""
                    $Status = "Supported"
                    if ($CompliancePolicy.deviceThreatProtectionEnabled) {
                        $Status = "Unsupported"
                        $PolicyErrors++
                        $SettingValue = $CompliancePolicy.advancedThreatProtectionRequiredSecurityLevel
                        $Comment = $URLSupportedCompliancePolicies
                    }
                    $SettingPSObj = [PSCustomObject][Ordered]@{
                        ID                    = $ID
                        PolicyName            = $CompliancePolicy.displayName
                        PolicyID              = $CompliancePolicy.id
                        PolicyType            = $CPType
                        AssignedToGroup       = $outAssignedToGroup
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        ExcludedFromGroupList = $ExcludedFromGroup
                        TeamsDevicesStatus    = $Status 
                        Setting               = $Setting
                        Value                 = $SettingValue
                        SettingDescription    = $SettingDescription
                        Comment               = $Comment
                    }
                    [void]$output.Add($SettingPSObj)
                    #endregion

                    #region 2: Device Health > Device managed with device administrator
                    $ID = 2
                    $Setting = "securityBlockDeviceAdministratorManagedDevices"
                    $SettingDescription = "Device Health > Device managed with device administrator"
                    $SettingValue = "Not Configured"
                    $Comment = ""
                    $Status = "Supported"
                    if ($CompliancePolicy.securityBlockDeviceAdministratorManagedDevices) {
                        $Status = "Unsupported"
                        $SettingValue = "Block"
                        $Comment = "Teams Android devices management requires device administrator to be enabled."
                        $PolicyErrors++
                    }
                    $SettingPSObj = [PSCustomObject][Ordered]@{
                        ID                    = $ID
                        PolicyName            = $CompliancePolicy.displayName
                        PolicyID              = $CompliancePolicy.id
                        PolicyType            = $CPType
                        AssignedToGroup       = $outAssignedToGroup
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        ExcludedFromGroupList = $ExcludedFromGroup
                        TeamsDevicesStatus    = $Status 
                        Setting               = $Setting
                        Value                 = $SettingValue
                        SettingDescription    = $SettingDescription
                        Comment               = $Comment
                    }
                    [void]$output.Add($SettingPSObj)
                    #endregion

                    #region 4: Device Health > Require the device to be at or under the Device Threat Level
                    $ID = 4
                    $Setting = "deviceThreatProtectionRequiredSecurityLevel"
                    $SettingDescription = "Device Health > Require the device to be at or under the Device Threat Level"
                    $SettingValue = "Not Configured"
                    $Comment = ""
                    $Status = "Supported"
                    if ($CompliancePolicy.deviceThreatProtectionRequiredSecurityLevel -ne "unavailable") {
                        $Status = "Unsupported"
                        $SettingValue = $CompliancePolicy.deviceThreatProtectionRequiredSecurityLevel
                        $Comment = $URLSupportedCompliancePolicies
                        $PolicyErrors++
                    }
                    $SettingPSObj = [PSCustomObject][Ordered]@{
                        ID                    = $ID
                        PolicyName            = $CompliancePolicy.displayName
                        PolicyID              = $CompliancePolicy.id
                        PolicyType            = $CPType
                        AssignedToGroup       = $outAssignedToGroup
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        ExcludedFromGroupList = $ExcludedFromGroup
                        TeamsDevicesStatus    = $Status 
                        Setting               = $Setting
                        Value                 = $SettingValue
                        SettingDescription    = $SettingDescription
                        Comment               = $Comment
                    }
                    [void]$output.Add($SettingPSObj)
                    #endregion

                    #region 5: Device Health > Google Protect > Google Play Services is Configured
                    $ID = 5
                    $Setting = "securityRequireGooglePlayServices"
                    $SettingDescription = "Device Health > Google Protect > Google Play Services is Configured"
                    $SettingValue = "Not Configured"
                    $Comment = ""
                    $Status = "Supported"
                    if ($CompliancePolicy.securityRequireGooglePlayServices) {
                        $Status = "Unsupported"
                        $SettingValue = "Require"
                        $Comment = "Google play isn't installed on Teams Android devices."
                        $PolicyErrors++
                    }
                    $SettingPSObj = [PSCustomObject][Ordered]@{
                        ID                    = $ID
                        PolicyName            = $CompliancePolicy.displayName
                        PolicyID              = $CompliancePolicy.id
                        PolicyType            = $CPType
                        AssignedToGroup       = $outAssignedToGroup
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        ExcludedFromGroupList = $ExcludedFromGroup
                        TeamsDevicesStatus    = $Status 
                        Setting               = $Setting
                        Value                 = $SettingValue
                        SettingDescription    = $SettingDescription
                        Comment               = $Comment
                    }
                    [void]$output.Add($SettingPSObj)
                    #endregion

                    #region 6: Device Health > Google Protect > Up-to-date security provider
                    $ID = 6
                    $Setting = "securityRequireUpToDateSecurityProviders"
                    $SettingDescription = "Device Health > Google Protect > Up-to-date security provider"
                    $SettingValue = "Not Configured"
                    $Comment = ""
                    $Status = "Supported"
                    if ($CompliancePolicy.securityRequireUpToDateSecurityProviders) {
                        $Status = "Unsupported"
                        $SettingValue = "Require"
                        $Comment = "Google play isn't installed on Teams Android devices."
                        $PolicyErrors++
                    }
                    $SettingPSObj = [PSCustomObject][Ordered]@{
                        ID                    = $ID
                        PolicyName            = $CompliancePolicy.displayName
                        PolicyID              = $CompliancePolicy.id
                        PolicyType            = $CPType
                        AssignedToGroup       = $outAssignedToGroup
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        ExcludedFromGroupList = $ExcludedFromGroup
                        TeamsDevicesStatus    = $Status 
                        Setting               = $Setting
                        Value                 = $SettingValue
                        SettingDescription    = $SettingDescription
                        Comment               = $Comment
                    }
                    [void]$output.Add($SettingPSObj)
                    #endregion
                
                    #region 7: Device Health > Google Protect > Threat scan on apps
                    $ID = 7
                    $Setting = "securityRequireVerifyApps"
                    $SettingDescription = "Device Health > Google Protect > Threat scan on apps"
                    $SettingValue = "Not Configured"
                    $Comment = ""
                    $Status = "Supported"
                    if ($CompliancePolicy.securityRequireVerifyApps) {
                        $Status = "Unsupported"
                        $SettingValue = "Require"
                        $Comment = "Google play isn't installed on Teams Android devices."
                        $PolicyErrors++
                    }
                    $SettingPSObj = [PSCustomObject][Ordered]@{
                        ID                    = $ID
                        PolicyName            = $CompliancePolicy.displayName
                        PolicyID              = $CompliancePolicy.id
                        PolicyType            = $CPType
                        AssignedToGroup       = $outAssignedToGroup
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        ExcludedFromGroupList = $ExcludedFromGroup
                        TeamsDevicesStatus    = $Status 
                        Setting               = $Setting
                        Value                 = $SettingValue
                        SettingDescription    = $SettingDescription
                        Comment               = $Comment
                    }
                    [void]$output.Add($SettingPSObj)
                    #endregion

                    #region 8: Device Health > Google Protect > SafetyNet device attestation
                    $ID = 8
                    $Setting = "securityRequireSafetyNetAttestation"
                    $SettingDescription = "Device Health > Google Protect > SafetyNet device attestation"
                    $SettingValue = "Not Configured"
                    $Comment = ""
                    $Status = "Supported"
                    if (($CompliancePolicy.securityRequireSafetyNetAttestationBasicIntegrity) -or ($CompliancePolicy.securityRequireSafetyNetAttestationCertifiedDevice)) {
                        $Status = "Unsupported"
                        $Comment = "Google play isn't installed on Teams Android devices."
                        $PolicyErrors++
                        if ($CompliancePolicy.securityRequireSafetyNetAttestationCertifiedDevice) {
                            $SettingValue = "Check basic integrity and certified devices"
                        }
                        elseif ($CompliancePolicy.securityRequireSafetyNetAttestationBasicIntegrity) {
                            $SettingValue = "Check basic integrity"
                        }
                    }
                    $SettingPSObj = [PSCustomObject][Ordered]@{
                        ID                    = $ID
                        PolicyName            = $CompliancePolicy.displayName
                        PolicyID              = $CompliancePolicy.id
                        PolicyType            = $CPType
                        AssignedToGroup       = $outAssignedToGroup
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        ExcludedFromGroupList = $ExcludedFromGroup
                        TeamsDevicesStatus    = $Status 
                        Setting               = $Setting
                        Value                 = $SettingValue
                        SettingDescription    = $SettingDescription
                        Comment               = $Comment
                    }
                    [void]$output.Add($SettingPSObj)
                    #endregion

                    #region 11: System Security > Device Security > Block apps from unknown sources
                    $ID = 11
                    $Setting = "securityPreventInstallAppsFromUnknownSources"
                    $SettingDescription = "System Security > Device Security > Block apps from unknown sources"
                    $SettingValue = "Not Configured"
                    $Comment = ""
                    $Status = "Supported"
                    if ($CompliancePolicy.securityPreventInstallAppsFromUnknownSources) {
                        $Status = "Unsupported"
                        $SettingValue = "Block"
                        $Comment = "Only Teams admins install apps or OEM tools"
                        $PolicyErrors++
                    }
                    $SettingPSObj = [PSCustomObject][Ordered]@{
                        ID                    = $ID
                        PolicyName            = $CompliancePolicy.displayName
                        PolicyID              = $CompliancePolicy.id
                        PolicyType            = $CPType
                        AssignedToGroup       = $outAssignedToGroup
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        ExcludedFromGroupList = $ExcludedFromGroup
                        TeamsDevicesStatus    = $Status 
                        Setting               = $Setting
                        Value                 = $SettingValue
                        SettingDescription    = $SettingDescription
                        Comment               = $Comment
                    }
                    [void]$output.Add($SettingPSObj)
                    #endregion

                    #region 15: System Security > Device Security > Restricted apps
                    $ID = 15
                    $Setting = "securityPreventInstallAppsFromUnknownSources"
                    $SettingDescription = "System Security > Device Security > Restricted apps"
                    $SettingValue = "Not Configured"
                    $Comment = ""
                    $Status = "Supported"
                    if (($CompliancePolicy.restrictedApps).count -gt 0 ) {
                        $Status = "Unsupported"
                        $SettingValue = "Found " + ($CompliancePolicy.restrictedApps).count + " restricted app(s)"
                        $Comment = $URLSupportedCompliancePolicies
                        $PolicyErrors++
                    }
                    $SettingPSObj = [PSCustomObject][Ordered]@{
                        ID                    = $ID
                        PolicyName            = $CompliancePolicy.displayName
                        PolicyID              = $CompliancePolicy.id
                        PolicyType            = $CPType
                        AssignedToGroup       = $outAssignedToGroup
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        ExcludedFromGroupList = $ExcludedFromGroup
                        TeamsDevicesStatus    = $Status 
                        Setting               = $Setting
                        Value                 = $SettingValue
                        SettingDescription    = $SettingDescription
                        Comment               = $Comment
                    }
                    [void]$output.Add($SettingPSObj)
                    #endregion
                }
                        
                if ($CompliancePolicy."@odata.type" -in $SupportedAndroidCompliancePolicies) {
                    #region 3: Device Health > Rooted devices
                    $ID = 3
                    $Setting = "securityBlockJailbrokenDevices"
                    $SettingDescription = "Device Health > Rooted devices"
                    $SettingValue = "Not Configured"
                    $Comment = ""
                    $Status = "Supported"
                    if ($CompliancePolicy.securityBlockJailbrokenDevices) {
                        $Status = "Warning"
                        $SettingValue = "Block"
                        $Comment = "This setting can cause sign in issues."
                        $PolicyWarnings++
                    }
                    $SettingPSObj = [PSCustomObject][Ordered]@{
                        ID                    = $ID
                        PolicyName            = $CompliancePolicy.displayName
                        PolicyID              = $CompliancePolicy.id
                        PolicyType            = $CPType
                        AssignedToGroup       = $outAssignedToGroup
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        ExcludedFromGroupList = $ExcludedFromGroup
                        TeamsDevicesStatus    = $Status 
                        Setting               = $Setting
                        Value                 = $SettingValue
                        SettingDescription    = $SettingDescription
                        Comment               = $Comment
                    }
                    [void]$output.Add($SettingPSObj)
                    #endregion

                    #region 10: System Security > Encryption > Require encryption of data storage on device.
                    $ID = 10
                    $Setting = "storageRequireEncryption"
                    $SettingDescription = "System Security > Encryption > Require encryption of data storage on device"
                    $SettingValue = "Not Configured"
                    $Comment = ""
                    $Status = "Supported"
                    if ($CompliancePolicy.storageRequireEncryption) {
                        $Status = "Warning"
                        $SettingValue = "Require"
                        $Comment = "Manufacturers might configure encryption attributes on their devices in a way that Intune doesn't recognize. If this happens, Intune marks the device as noncompliant."
                        $PolicyWarnings++
                    }
                    $SettingPSObj = [PSCustomObject][Ordered]@{
                        ID                    = $ID
                        PolicyName            = $CompliancePolicy.displayName
                        PolicyID              = $CompliancePolicy.id
                        PolicyType            = $CPType
                        AssignedToGroup       = $outAssignedToGroup
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        ExcludedFromGroupList = $ExcludedFromGroup
                        TeamsDevicesStatus    = $Status 
                        Setting               = $Setting
                        Value                 = $SettingValue
                        SettingDescription    = $SettingDescription
                        Comment               = $Comment
                    }
                    [void]$output.Add($SettingPSObj)
                    #endregion
                            
                    #region 14: System Security > Device Security > Minimum security patch level
                    $ID = 14
                    $Setting = "minAndroidSecurityPatchLevel"
                    $SettingDescription = "System Security > Device Security > Minimum security patch level"
                    $SettingValue = "Not Configured"
                    $Comment = ""
                    $Status = "Supported"
                    if (!([string]::IsNullOrEmpty($CompliancePolicy.minAndroidSecurityPatchLevel))) {
                        $Status = "Warning"
                        $SettingValue = $CompliancePolicy.minAndroidSecurityPatchLevel
                        $Comment = "This setting can cause sign in issues."
                        $PolicyWarnings++
                    }
                    $SettingPSObj = [PSCustomObject][Ordered]@{
                        ID                    = $ID
                        PolicyName            = $CompliancePolicy.displayName
                        PolicyID              = $CompliancePolicy.id
                        PolicyType            = $CPType
                        AssignedToGroup       = $outAssignedToGroup
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        ExcludedFromGroupList = $ExcludedFromGroup
                        TeamsDevicesStatus    = $Status 
                        Setting               = $Setting
                        Value                 = $SettingValue
                        SettingDescription    = $SettingDescription
                        Comment               = $Comment
                    }
                    [void]$output.Add($SettingPSObj)
                    #endregion

                    #region 16: System Security > All Android devices > Maximum minutes of inactivity before password is required
                    $ID = 16
                    $Setting = "passwordMinutesOfInactivityBeforeLock"
                    $SettingDescription = "System Security > All Android devices > Maximum minutes of inactivity before password is required"
                    $SettingValue = "Not Configured"
                    $Comment = ""
                    $Status = "Supported"
                    if (!([string]::IsNullOrEmpty($CompliancePolicy.passwordMinutesOfInactivityBeforeLock))) {
                        $Status = "Unsupported"
                        $SettingValue = "" + $CompliancePolicy.passwordMinutesOfInactivityBeforeLock + " minutes"
                        $Comment = $URLSupportedCompliancePolicies
                        $PolicyErrors++
                    }
                    $SettingPSObj = [PSCustomObject][Ordered]@{
                        ID                    = $ID
                        PolicyName            = $CompliancePolicy.displayName
                        PolicyID              = $CompliancePolicy.id
                        PolicyType            = $CPType
                        AssignedToGroup       = $outAssignedToGroup
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        ExcludedFromGroupList = $ExcludedFromGroup
                        TeamsDevicesStatus    = $Status 
                        Setting               = $Setting
                        Value                 = $SettingValue
                        SettingDescription    = $SettingDescription
                        Comment               = $Comment
                    }
                    [void]$output.Add($SettingPSObj)
                    #endregion
                } 
                elseif ($CompliancePolicy."@odata.type" -in $SupportedWindowsCompliancePolicies) {

                    #region 18: Device Properties > Operation System Version
                    $ID = 18.1
                    $Setting = "mobileOsMinimumVersion"
                    $SettingDescription = "Device Properties > Operation System Version > Minimum OS version for mobile devices"
                    $SettingValue = "Not Configured"
                    $Comment = ""
                    $Status = "Supported"
                    if (!([string]::IsNullOrEmpty($CompliancePolicy.mobileOsMinimumVersion))) {
                        $Status = "Unsupported"
                        $SettingValue = $CompliancePolicy.mobileOsMinimumVersion
                        $Comment = $URLSupportedCompliancePolicies
                        $PolicyErrors++
                    }
                    $SettingPSObj = [PSCustomObject][Ordered]@{
                        ID                    = $ID
                        PolicyName            = $CompliancePolicy.displayName
                        PolicyID              = $CompliancePolicy.id
                        PolicyType            = $CPType
                        AssignedToGroup       = $outAssignedToGroup
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        ExcludedFromGroupList = $ExcludedFromGroup
                        TeamsDevicesStatus    = $Status 
                        Setting               = $Setting
                        Value                 = $SettingValue
                        SettingDescription    = $SettingDescription
                        Comment               = $Comment
                    }
                    [void]$output.Add($SettingPSObj)
                
                    $ID = 18.2
                    $Setting = "mobileOsMaximumVersion"
                    $SettingDescription = "Device Properties > Operation System Version > Maximum OS version for mobile devices"
                    $SettingValue = "Not Configured"
                    $Status = "Supported"
                    $Comment = ""
                    if (!([string]::IsNullOrEmpty($CompliancePolicy.mobileOsMaximumVersion))) {
                        $Status = "Unsupported"
                        $SettingValue = $CompliancePolicy.mobileOsMaximumVersion
                        $Comment = $URLSupportedCompliancePolicies
                        $PolicyErrors++
                    }
                    $SettingPSObj = [PSCustomObject][Ordered]@{
                        ID                    = $ID
                        PolicyName            = $CompliancePolicy.displayName
                        PolicyID              = $CompliancePolicy.id
                        PolicyType            = $CPType
                        AssignedToGroup       = $outAssignedToGroup
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        ExcludedFromGroupList = $ExcludedFromGroup
                        TeamsDevicesStatus    = $Status 
                        Setting               = $Setting
                        Value                 = $SettingValue
                        SettingDescription    = $SettingDescription
                        Comment               = $Comment
                    }
                    [void]$output.Add($SettingPSObj)
                    #endregion

                    #region 19: Device Properties > Operation System Version > Valid operating system builds
                    $ID = 19
                    $Setting = "validOperatingSystemBuildRanges"
                    $SettingDescription = "Device Properties > Operation System Version > Valid operating system builds"
                    $SettingValue = "Not Configured"
                    $Status = "Supported"
                    $Comment = ""
                    if (!([string]::IsNullOrEmpty($CompliancePolicy.validOperatingSystemBuildRanges))) {
                        $Status = "Unsupported"
                        $SettingValue = "Found " + ($CompliancePolicy.validOperatingSystemBuildRanges).count + " valid OS configured build(s)"
                        $Comment = $URLSupportedCompliancePolicies
                        $PolicyErrors++
                    }
                    $SettingPSObj = [PSCustomObject][Ordered]@{
                        ID                    = $ID
                        PolicyName            = $CompliancePolicy.displayName
                        PolicyID              = $CompliancePolicy.id
                        PolicyType            = $CPType
                        AssignedToGroup       = $outAssignedToGroup
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        ExcludedFromGroupList = $ExcludedFromGroup
                        TeamsDevicesStatus    = $Status 
                        Setting               = $Setting
                        Value                 = $SettingValue
                        SettingDescription    = $SettingDescription
                        Comment               = $Comment
                    }
                    [void]$output.Add($SettingPSObj)
                    #endregion

                    #region 20: System Security > Defender > Microsoft Defender Antimalware minimum version
                    $ID = 20
                    $Setting = "defenderVersion"
                    $SettingDescription = "System Security > Defender > Microsoft Defender Antimalware minimum version"
                    $SettingValue = "Not Configured"
                    $Status = "Supported"
                    $Comment = ""
                    if (!([string]::IsNullOrEmpty($CompliancePolicy.defenderVersion))) {
                        $Status = "Unsupported"
                        $SettingValue = $CompliancePolicy.defenderVersion
                        $Comment = "Teams Rooms automatically updates this component so there's no need to set compliance policies."
                        $PolicyErrors++
                    }
                    $SettingPSObj = [PSCustomObject][Ordered]@{
                        ID                    = $ID
                        PolicyName            = $CompliancePolicy.displayName
                        PolicyID              = $CompliancePolicy.id
                        PolicyType            = $CPType
                        AssignedToGroup       = $outAssignedToGroup
                        AssignedToGroupList   = $AssignedToGroup
                        ExcludedFromGroup     = $outExcludedFromGroup 
                        ExcludedFromGroupList = $ExcludedFromGroup
                        TeamsDevicesStatus    = $Status 
                        Setting               = $Setting
                        Value                 = $SettingValue
                        SettingDescription    = $SettingDescription
                        Comment               = $Comment
                    }
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
                $PolicySum = [PSCustomObject][Ordered]@{
                    PolicyName            = $CompliancePolicy.displayName
                    PolicyID              = $CompliancePolicy.id
                    PolicyType            = $CPType
                    AssignedToGroup       = $outAssignedToGroup
                    AssignedToGroupList   = $AssignedToGroup
                    ExcludedFromGroup     = $outExcludedFromGroup 
                    ExcludedFromGroupList = $ExcludedFromGroup
                    TeamsDevicesStatus    = $StatusSum
                }
                $PolicySum.PSObject.TypeNames.Insert(0, 'TeamsDeviceCompliancePolicy')
                [void]$outputSum.Add($PolicySum) 
            }
            elseif (($AssignedToGroup.count -eq 0) -and !($UserUPN -or $DeviceID -or $Detailed)) {
                $skippedCompliancePolicies++
            }
        }
    }
    #region Output
    #2025-03-28: Improving the output code readability 
    if ($totalCompliancePolicies -eq 0) {
        if ($UserUPN) {
            Write-Warning ("The user " + $UserUPN + " doesn't have any Compliance Policies assigned.")
        }
        else {
            Write-Warning "No Compliance Policies assigned to All Users, All Devices or group found. Please use Test-UcTeamsDevicesCompliancePolicy -All to check all policies."
        }
        return
    }
    if (!$IncludeSupported) {
        $output = $output | Where-Object -Property TeamsDevicesStatus -NE -Value "Supported"
    }
    if ($Detailed) {
        if ($output.count -eq 0 -and !$IncludeSupported) {
            Write-Warning "No unsupported settings found, please use Test-UcTeamsDevicesCompliancePolicy -IncludeSupported to output all settings."
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
        if (($skippedCompliancePolicies -gt 0) -and !$All) {
            Write-Warning ("Skipping $skippedCompliancePolicies compliance policies since they will not be applied to Teams Devices.")
            Write-Warning ("Please use the All switch to check all policies: Test-UcTeamsDevicesCompliancePolicy -All")
        }
        if ($displayWarning) {
            Write-Warning "One or more policies contain unsupported settings, please use Test-UcTeamsDevicesCompliancePolicy -Detailed to identify the unsupported settings."
        }
        return $outputSum | Sort-Object PolicyName            
    }
    #endregion
}

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

function Test-UcTeamsDevicesEnrollmentProfile {
    <#
        .SYNOPSIS
        Validate Intune Enrollment Profiles that are supported by Microsoft Teams Android Devices

        .DESCRIPTION
        This function will validate if the settings in a Enrollment Profile are supported:

        Contributors: David Paulino, Gonçalo Sepúlveda and Eileen Beato

        Requirements:   EntraAuth PowerShell Module (Install-Module EntraAuth)
                        or
                        Microsoft Graph Authentication PowerShell Module (Install-Module Microsoft.Graph.Authentication)

                        Microsoft Graph Scopes:
                            "DeviceManagementServiceConfig.Read.All", "DeviceManagementConfiguration.Read.All", "Directory.Read.All"

        .PARAMETER UserUPN
        Specifies a UserUPN that we want to check for a user enrollment profiles.

        .PARAMETER PlatformType
        Platform Type:
            AndroidDeviceAdministrator
            AOSP


        .PARAMETER Detailed
        Displays test results for unsupported settings in each Intune Enrollment Profile.

        .PARAMETER ExportCSV
        When present will export the detailed results to a CSV file. By defautl will save the file under the current user downloads, unless we specify the OutputPath.

        .PARAMETER OutputPath
        Allows to specify the path where we want to save the results.

        .EXAMPLE 
        PS> Test-UcTeamsDevicesEnrollmentProfile

        .EXAMPLE 
        PS> Test-UcTeamsDevicesEnrollmentProfile -UserUPN
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [string]$UserUPN,
        [ValidateSet("AndroidDeviceAdministrator", "AOSP", "Any")]
        [string]$PlatformType = "Any",
        [switch]$Detailed,
        [switch]$ExportCSV,
        [string]$OutputPath
    )   
    
    $GraphPathUsers = "/users"
    $GraphPathGroups = "/groups"
    $GraphPathEnrollmentProfiles = "/deviceManagement/deviceEnrollmentConfigurations/?`$expand=assignments"
    $GraphPathAOSPEnrollmentProfiles = "/deviceManagement/androidDeviceOwnerEnrollmentProfiles?`$select=displayName,tokenExpirationDateTime&`$filter=enrollmentMode eq 'corporateOwnedAOSPUserAssociatedDevice' and isTeamsDeviceProfile eq true" 
    $output = [System.Collections.ArrayList]::new()

    #region Graph Connection, Scope validation and module version
    if (!(Test-UcServiceConnection -Type MSGraph -Scopes "DeviceManagementServiceConfig.Read.All", "DeviceManagementConfiguration.Read.All", "Directory.Read.All")) {
        return
    }
    #2025-01-31: Only need to check this once per PowerShell session
    if (!($global:UCLobbyTeamsModuleCheck)) {
        Test-UcPowerShellModule -ModuleName UcLobbyTeams | Out-Null
        $global:UCLobbyTeamsModuleCheck = $true
    }
    #endregion

    #region ExportCSV defaults to Detaileld and Output Path Test
    if ($ExportCSV) {
        $Detailed = $true
        $outFileName = "TeamsDevices_EnrollmentProfile_Report_" + ( get-date ).ToString('yyyyMMdd-HHmmss') + ".csv"
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
        Write-Progress -Activity "Test-UcTeamsDevicesEnrollmentProfile" -Status "Getting Intune Enrollment Policies"
        $EnrollmentProfiles = Invoke-UcGraphRequest -Path $GraphPathEnrollmentProfiles -Beta
        if ($PlatformType -in ("Any", "AOSP") ) {
            Write-Progress -Activity "Test-UcTeamsDevicesEnrollmentProfile" -Status "Getting AOSP Enrollment Profiles"
            $AOSPEnrollmentProfiles = Invoke-UcGraphRequest -Path $GraphPathAOSPEnrollmentProfiles -Beta
        }
    }
    catch [System.Net.Http.HttpRequestException] {
        if ($PSItem.Exception.Response.StatusCode -eq "Unauthorized") {
            Write-Error "Access Denied, please make sure the user connecing to MS Graph is part of one of the following Global Reader/Intune Service Administrator/Global Administrator roles"
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
            Write-Progress -Activity "Test-UcTeamsDevicesEnrollmentProfile" -Status "Retriving User Group Membership"
            $UserGroups = (Invoke-UcGraphRequest -Path ($GraphPathUsers + "/" + $UserUPN + "/transitiveMemberOf?`$select=id")).id
        }
        catch [System.Net.Http.HttpRequestException] {
            if ($PSItem.Exception.Response.StatusCode -eq "NotFound") {
                Write-warning -Message ("User Not Found: " + $UserUPN)
            }
            return
        }
    }
                 
    foreach ($EnrollmentProfile in $EnrollmentProfiles) {
        $Status = "Not Supported"
        #This is the default enrollment policy that is applied to all users/devices
        if ($EnrollmentProfile."@odata.type" -eq "#microsoft.graph.deviceEnrollmentPlatformRestrictionsConfiguration") {
            if ($PlatformType -in ("Any", "AndroidDeviceAdministrator")) {
                if (!($EnrollmentProfile.androidRestriction.platformBlocked) -and !($EnrollmentProfile.androidRestriction.personalDeviceEnrollmentBlocked)) {
                    $Status = "Supported"
                }
                $SettingPSObj = [PSCustomObject][Ordered]@{
                    ID                              = 9999
                    EnrollmentProfileName           = $EnrollmentProfile.displayName
                    Priority                        = "Default"
                    PlatformType                    = "Android device administrator"
                    AssignedToGroup                 = "All devices"
                    TeamsDevicesStatus              = $Status
                    PlatformBlocked                 = $EnrollmentProfile.androidRestriction.platformBlocked
                    PersonalDeviceEnrollmentBlocked = $EnrollmentProfile.androidRestriction.personalDeviceEnrollmentBlocked
                    osMinimumVersion                = $EnrollmentProfile.androidRestriction.osMinimumVersion
                    osMaximumVersion                = $EnrollmentProfile.androidRestriction.osMaximumVersion
                    blockedManufacturers            = $EnrollmentProfile.androidRestriction.blockedManufacturers
                    TokenExpirationDate             = ""
                }
                $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceEnrollmentProfile')
                [void]$output.Add($SettingPSObj)
            }
        }

        $Status = "Not Supported"
        if (($EnrollmentProfile."@odata.type" -eq "#microsoft.graph.deviceEnrollmentPlatformRestrictionConfiguration") -and ($EnrollmentProfile.platformType -eq "android") ) {
            $AssignedToGroup = [System.Collections.ArrayList]::new()

            if ($UserUPN -and !($EnrollmentProfile.assignments.target.groupId | Where-Object { $_ -in $UserGroups })) {
                #Skipping this policy since not assigned to a user
                Write-Verbose ("Skipping Enrollment Profile " + $EnrollmentProfile.displayName + " because user $UserUPN is not part of the Included Groups.")
                continue
            }

            #2025-03-28: Moving the Group Name here since we use expand Assignments in the graph request.
            foreach ($AssignedGroupID in $EnrollmentProfile.assignments.target.groupId) {
                $GroupInfo = Invoke-UcGraphRequest -Path "$GraphPathGroups/$AssignedGroupID/?`$select=id,displayname"
                $GroupEntry = [PSCustomObject][Ordered]@{
                    GroupID          = $AssignedGroupID
                    GroupDisplayName = $GroupInfo.displayname
                }
                [void]$AssignedToGroup.Add($GroupEntry)
            }

            if ($AssignedToGroup.Count -gt 0) {
                $outAssignedToGroup = "None"
                if ($AssignedToGroup.count -eq 1) {
                    $outAssignedToGroup = $AssignedToGroup[0].GroupDisplayName
                }
                elseif ($AssignedToGroup.count -gt 1) {
                    $outAssignedToGroup = "" + $AssignedToGroup.count + " group(s)"
                }
                if (!($EnrollmentProfile.platformRestriction.platformBlocked) -and !($EnrollmentProfile.platformRestriction.personalDeviceEnrollmentBlocked)) {
                    $Status = "Supported"
                }
                $SettingPSObj = [PSCustomObject][Ordered]@{
                    ID                              = $EnrollmentProfile.priority
                    EnrollmentProfileName           = $EnrollmentProfile.displayName
                    Priority                        = $EnrollmentProfile.priority
                    PlatformType                    = "Android device administrator"
                    AssignedToGroup                 = $outAssignedToGroup
                    AssignedToGroupList             = $AssignedToGroup
                    TeamsDevicesStatus              = $Status 
                    PlatformBlocked                 = $EnrollmentProfile.platformRestriction.platformBlocked
                    PersonalDeviceEnrollmentBlocked = $EnrollmentProfile.platformRestriction.personalDeviceEnrollmentBlocked
                    osMinimumVersion                = $EnrollmentProfile.platformRestriction.osMinimumVersion
                    osMaximumVersion                = $EnrollmentProfile.platformRestriction.osMaximumVersion
                    blockedManufacturers            = $EnrollmentProfile.platformRestriction.blockedManufacturers
                    TokenExpirationDate             = ""
                }
                $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceEnrollmentProfile')
                [void]$output.Add($SettingPSObj)
            }
        }
    }
    #region 2024-09-12: Support for Android AOSP Enrollment
    if ($PlatformType -in ("Any", "AOSP") ) {
        $PolicyName = ""
        $TokenExpirationDate = ""
        $CurrentDate = Get-Date
        if ($AOSPEnrollmentProfiles.Length -ge 1) {
            #In some cases we can have multiple AOSP profiles enabled for Teams Devices, currently we only support one valid at a time.
            $ValidAOSPEnrollmentProfiles = $AOSPEnrollmentProfiles | Where-Object { $_.tokenExpirationDateTime -gt $CurrentDate }
            if ($ValidAOSPEnrollmentProfiles.Length -eq 1) {
                $PolicyName = $ValidAOSPEnrollmentProfiles.displayName
                $TokenExpirationDate = $ValidAOSPEnrollmentProfiles.tokenExpirationDateTime
                $TeamsDevicesStatus = "Supported - Token valid for " + ($TokenExpirationDate - $CurrentDate).days + " day(s)"
            }
            elseif ($ValidAOSPEnrollmentProfiles.Length -eq 0) {
                $ExpiredAOSPEnrollmentProfiles = $AOSPEnrollmentProfiles | Sort-Object -Property tokenExpirationDateTime -Descending
                $TeamsDevicesStatus = "Not Supported - Token Expired on " + ($ExpiredAOSPEnrollmentProfiles[0].tokenExpirationDateTime.ToString([cultureinfo]::CurrentCulture.DateTimeFormat.ShortDatePattern))
            }
            else {
                $TeamsDevicesStatus = "Not Supported - Multiple AOSP Enrollment Profile enabled for Teams Devices"
            }
        }
        else {
            $TeamsDevicesStatus = "Not Supported - Missing AOSP Enrollment"
        }
        $SettingPSObj = [PSCustomObject][Ordered]@{
            ID                              = "10000"
            EnrollmentProfileName           = $PolicyName
            Priority                        = "ASOP"
            PlatformType                    = "Android Open Source Project (AOSP)"
            TeamsDevicesStatus              = $TeamsDevicesStatus
            PlatformBlocked                 = ""
            PersonalDeviceEnrollmentBlocked = ""
            osMinimumVersion                = ""
            osMaximumVersion                = ""
            blockedManufacturers            = ""
            TokenExpirationDate             = $TokenExpirationDate
        }
        $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceEnrollmentProfile')
        [void]$output.Add($SettingPSObj)
    }
    #endregion
    #region Output
    if ($ExportCSV) {
        $output | Sort-Object ID | Select-Object -ExcludeProperty ID | Export-Csv -path $OutputFullPath -NoTypeInformation
        Write-Host ("Results available in: " + $OutputFullPath) -ForegroundColor Cyan
        return
    }
    if ($Detailed) {
        return $output | Sort-Object ID | Select-Object -ExcludeProperty ID 
    }
    else {
        return $output | Sort-Object ID
    }
    #endregion
}

function Test-UcTeamsOnlyDNSRequirements {
    <#
        .SYNOPSIS
        Check if the DNS records are OK for a TeamsOnly Tenant.

        .DESCRIPTION
        This function will check if the DNS entries that were previously required.

        .PARAMETER Domain
        Specifies a domain registered with Microsoft 365

        .EXAMPLE
        PS> Test-UcTeamsOnlyDNSRequirements

        .EXAMPLE
        PS> Test-UcTeamsOnlyDNSRequirements -Domain uclobby.com
    #>
    param(
        [Parameter(Mandatory = $false)]
        [string]$Domain,
        [switch]$All
    )

    $outDNSRecords = [System.Collections.ArrayList]::new()
    if ($Domain) {
        $365Domains = Get-UcM365Domains -Domain $Domain | Where-Object { $_.Name -notlike "*.onmicrosoft.com" }
    }
    else {
        try {
            #2025-01-31: Only need to check this once per PowerShell session
            if (!($global:UCLobbyTeamsModuleCheck)) {
                Test-UcPowerShellModule -ModuleName UcLobbyTeams | Out-Null
                $global:UCLobbyTeamsModuleCheck = $true
            }
            #We only need to validate the Enable domains and exclude *.onmicrosoft.com
            $365Domains = Get-CsOnlineSipDomain | Where-Object { $_.Status -eq "Enabled" -and $_.Name -notlike "*.onmicrosoft.com" }
        }
        catch {
            Write-Host "Error: Please connect to Microsoft Teams PowerShell before running this cmdlet with Connect-MicrosoftTeams" -ForegroundColor Red
            return
        }
    }
    $DomainCount = ($365Domains.Count)
    $i = 1
    foreach ($365Domain in $365Domains) {
        $tmpDomain = $365Domain.Name
        Write-Progress -Activity "Teams Only DNS Requirements" -Status "Checking: $tmpDomain - $i of $DomainCount"
        $i++
        $DiscoverFQDN = "lyncdiscover." + $365Domain.Name
        $SIPFQDN = "sip." + $365Domain.Name
        $SIPTLSFQDN = "_sip._tls." + $365Domain.Name
        $FederationFQDN = "_sipfederationtls._tcp." + $365Domain.Name
        $DNSResultDiscover = (Resolve-DnsName $DiscoverFQDN -Type CNAME -ErrorAction Ignore).NameHost
        $DNSResultSIP = (Resolve-DnsName $SIPFQDN -Type CNAME -ErrorAction Ignore).NameHost
        $DNSResultSIPTLS = (Resolve-DnsName $SIPTLSFQDN -Type SRV -ErrorAction Ignore).NameTarget
        $DNSResultFederation = (Resolve-DnsName $FederationFQDN -Type SRV -ErrorAction Ignore)

        $DNSDiscover = ""
        if ($DNSResultDiscover -and ($All -or $DNSResultDiscover.contains("online.lync.com") -or $DNSResultDiscover.contains("online.gov.skypeforbusiness.us"))) {
            $DNSDiscover = $DNSResultDiscover
        }

        $DNSSIP = ""
        if ($DNSResultSIP -and ($All -or $DNSResultSIP.contains("online.lync.com") -or $DNSResultSIP.contains("online.gov.skypeforbusiness.us"))) {
            $DNSSIP = $DNSResultSIP
        }
        
        $DNSSIPTLS = ""
        if ($DNSResultSIPTLS -and ($All -or $DNSResultSIPTLS.contains("online.lync.com") -or $DNSResultSIPTLS.contains("online.gov.skypeforbusiness.us"))) {
            $DNSSIPTLS = $DNSResultSIPTLS
        }

        $DNSFederation = ""
        if ([string]::IsNullOrEmpty($DNSResultFederation.NameTarget)) {
            $DNSFederation = "Not configured"
        }
        elseif (($DNSResultFederation.NameTarget.equals("sipfed.online.lync.com") -or $DNSResultFederation.NameTarget.equals("sipfed.online.gov.skypeforbusiness.us")) -and $DNSResultFederation.Port -eq 5061) {
            $DNSFederation = "OK"
        }
        else {
            $DNSFederation = "NOK - " + $DNSResultFederation.NameTarget + ":" + $DNSResultFederation.Port 
        }

        if ($DNSDiscover -or $DNSSIP -or $DNSSIPTLS -or $DNSFederation) {
            $tmpDNSRecord = New-Object -TypeName PSObject -Property @{
                Domain           = $365Domain.Name
                DiscoverRecord   = $DNSDiscover
                SIPRecord        = $DNSSIP
                SIPTLSRecord     = $DNSSIPTLS
                FederationRecord = $DNSFederation
            }
            $tmpDNSRecord.PSObject.TypeNames.Insert(0, 'TeamsOnlyDNSRequirements')
            [void]$outDNSRecords.Add($tmpDNSRecord)
        }
    }
    return $outDNSRecords | Sort-Object Domain
}

function Update-UcTeamsDevice {
    <#
        .SYNOPSIS
        Update Microsoft Teams Devices

        .DESCRIPTION
        This function will send update commands to Teams Android Devices using MS Graph.

        Contributors: Eileen Beato, David Paulino and Bryan Kendrick

        Requirements:   EntraAuth PowerShell Module (Install-Module EntraAuth)
                        or
                        Microsoft Graph Authentication PowerShell Module (Install-Module Microsoft.Graph.Authentication)

                        Microsoft Graph Permissions:
                            "TeamworkDevice.ReadWrite.All","User.Read.All"

        .PARAMETER DeviceID
        Specify the Teams Admin Center Device ID that we want to update.

        .PARAMETER DeviceType
        Specifies a filter, valid options:
            Phone - Teams Native Phones
            MTRA - Microsoft Teams Room Running Android
            Display - Microsoft Teams Displays 
            Panel - Microsoft Teams Panels

        .PARAMETER UpdateType
        Allow to specify which time of update we want to do:
            Firmware
            TeamsClient
            All

        .PARAMETER SoftwareVersion
        Allow to specify which version we want to update.

        .PARAMETER InputCSV
        When present will use this file as Input, we only need a column with Device Id. It supports files exported from Teams Admin Center (TAC).

        .PARAMETER Subnet
        Only available when using InputCSV and requires a “IP Address” column, it allows to only send updates to Teams Android devices within a subnet.
        Format examples:
            10.0.0.0/8
            192.168.0.0/24

        .PARAMETER OutputPath
        Allows to specify the path where we want to save the results. By default will save on current user Download.

        .PARAMETER ReportOnly
        Will read Teams Device Android versions info and generate a report

        .EXAMPLE 
        PS> Update-UcTeamsDevice

        .EXAMPLE 
        PS> Update-UcTeamsDevice -ReportOnly

        .EXAMPLE
        PS> Update-UcTeamsDevice -InputCSV C:\Temp\DevicesList_2023-04-20_15-19-00-UTC.csv
    #>
    [cmdletbinding(SupportsShouldProcess)]
    param(
        [string]$DeviceID,
        [ValidateSet("Firmware", "TeamsClient", "All")]
        [string]$DeviceType,
        [string]$UpdateType = "All",
        [ValidateSet("Phone", "MTRA", "Display", "Panel")]       
        [string]$SoftwareVersion,
        [string]$InputCSV,
        [string]$Subnet,
        [string]$OutputPath,
        [switch]$ReportOnly
    )
    
    $regExIPAddressSubnet = "^((25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9])\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[0-9]))\/(3[0-2]|[1-2]{1}[0-9]{1}|[1-9])$"
    $outTeamsDevices = [System.Collections.ArrayList]::new()
    
    #2024-10-23: If report Only then we dont need write permission on TeamworkDevice
    $GraphConnected = $false
    if ($ReportOnly) {
        $GraphConnected = Test-UcServiceConnection -Type MSGraph -Scopes "TeamworkDevice.Read.All", "User.Read.All" -AltScopes "TeamworkDevice.Read.All", "User.ReadBasic.All"
    }
    else {
        $GraphConnected = Test-UcServiceConnection -Type MSGraph -Scopes "TeamworkDevice.ReadWrite.All", "User.Read.All" -AltScopes "TeamworkDevice.ReadWrite.All", "User.ReadBasic.All"
    }
    if (!($GraphConnected)) {
        return
    }
    
    #2025-01-31: Only need to check this once per PowerShell session
    if (!($global:UCLobbyTeamsModuleCheck)) {
        Test-UcPowerShellModule -ModuleName UcLobbyTeams | Out-Null
        $global:UCLobbyTeamsModuleCheck = $true
    }
    #Checking if the Subnet is valid
    if ($Subnet) {
        if (!($Subnet -match $regExIPAddressSubnet)) {
            Write-Host ("Error: Subnet " + $Subnet + " is invalid, please make sure the subnet is valid and in this format 10.0.0.0/8, 192.168.0.0/24") -ForegroundColor Red
            return
        } 
    }

    if ($ReportOnly) {
        $outFileName = "UpdateTeamsDevices_ReportOnly_" + ( get-date ).ToString('yyyyMMdd-HHmmss') + ".csv"
        $StatusType = "offline", "critical", "nonUrgent", "healthy"
    }
    else {
        $outFileName = "UpdateTeamsDevices_" + ( get-date ).ToString('yyyyMMdd-HHmmss') + ".csv"
        $StatusType = "critical", "nonUrgent"
    }
    #Verify if the Output Path exists
    if ($OutputPath) {
        if (!(Test-Path $OutputPath -PathType Container)) {
            Write-Host ("Error: Invalid folder " + $OutputPath) -ForegroundColor Red
            return
        }
        else {
            $OutputFullPath = [System.IO.Path]::Combine($OutputPath, $outFileName)
        }
    }
    else {                
        $OutputFullPath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads", $outFileName)
    }

    $graphRequests = [System.Collections.ArrayList]::new()
    if ($DeviceID) {
        $gRequestTmp = New-Object -TypeName PSObject -Property @{
            id     = $DeviceID
            method = "GET"
            url    = "/teamwork/devices/" + $DeviceID
        }
        [void]$graphRequests.Add($gRequestTmp)
        $GraphResponse = Invoke-UcGraphRequest -Requests $graphRequests Beta -Activity "Update-UcTeamsDevices, getting device info" -IncludeBody
            
        if ($GraphResponse.status -eq 200) {
            $TeamsDeviceList = $GraphResponse.body
        }
        elseif ($GraphResponse.status -eq 404) {
            Write-Host ("Error: Device ID $DeviceID not found.") -ForegroundColor Red
            return
        }
    }
    elseif ($InputCSV) {
        if (Test-Path $InputCSV) {
            try {
                $TeamsDeviceInput = Import-Csv -Path $InputCSV
            }
            catch {
                Write-Host ("Invalid CSV input file: " + $InputCSV) -ForegroundColor Red
                return
            }
            foreach ($TeamsDevice in $TeamsDeviceInput) {
                $includeDevice = $true
                if ($Subnet) {
                    $includeDevice = Test-UcIPaddressInSubnet -IPAddress $TeamsDevice.'IP Address' -Subnet $Subnet
                }

                if ($includeDevice) {
                    $gRequestTmp = New-Object -TypeName PSObject -Property @{
                        id     = $TeamsDevice.'Device Id'
                        method = "GET"
                        url    = "/teamwork/devices/" + $TeamsDevice.'Device Id'
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                }
            }
            if ($graphRequests.Count -gt 0) {
                $TeamsDeviceList = (Invoke-UcGraphRequest -Requests $graphRequests -Beta -Activity "Update-UcTeamsDevices, getting device info" )
            } 
        }
        else {
            Write-Host ("Error: File not found " + $InputCSV) -ForegroundColor Red
            return
        }
    }
    else {
        #Currently only Android based Teams devices are supported.
        switch ($DeviceType) {
            "Phone" { 
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "ipPhone"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'ipPhone'"
                }
                [void]$graphRequests.Add($gRequestTmp)
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "lowCostPhone"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'lowCostPhone'"
                }
                [void]$graphRequests.Add($gRequestTmp)
            }
            "MTRA" {            
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "collaborationBar"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'collaborationBar'"
                }
                [void]$graphRequests.Add($gRequestTmp)
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "touchConsole"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'touchConsole'"
                }
                [void]$graphRequests.Add($gRequestTmp)
            }
            "Display" {
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "teamsDisplay"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'teamsDisplay'"
                }
                [void]$graphRequests.Add($gRequestTmp)
            }
            "Panel" {
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "teamsPanel"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'teamsPanel'"
                }
                [void]$graphRequests.Add($gRequestTmp)
            }
            Default {
                #This is the only way to exclude MTRW and SurfaceHub by creating a request per device type.
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "ipPhone"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'ipPhone'"
                }
                [void]$graphRequests.Add($gRequestTmp) 
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "lowCostPhone"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'lowCostPhone'"
                }
                [void]$graphRequests.Add($gRequestTmp)
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "collaborationBar"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'collaborationBar'"
                }
                [void]$graphRequests.Add($gRequestTmp)
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "touchConsole"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'touchConsole'"
                }
                [void]$graphRequests.Add($gRequestTmp) 
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "teamsDisplay"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'teamsDisplay'"
                }
                [void]$graphRequests.Add($gRequestTmp)
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = "teamsPanel"
                    method = "GET"
                    url    = "/teamwork/devices/?`$filter=deviceType eq 'teamsPanel'"
                }
                [void]$graphRequests.Add($gRequestTmp)
            }
        }
        #Using new cmdlet to get a list of devices
        $TeamsDeviceList = (Invoke-UcGraphRequest -Requests $graphRequests -Beta -Activity "Update-UcTeamsDevices, getting device info").value
    }

    $devicesWithUpdatePending = 0
    $graphRequests = [System.Collections.ArrayList]::new()
    foreach ($TeamsDevice in $TeamsDeviceList) {
        if (($graphRequests.id -notcontains $TeamsDevice.currentuser.id) -and !([string]::IsNullOrEmpty($TeamsDevice.currentuser.id))) {
            $gRequestTmp = New-Object -TypeName PSObject -Property @{
                id     = $TeamsDevice.currentuser.id
                method = "GET"
                url    = "/users/" + $TeamsDevice.currentuser.id
            }
            [void]$graphRequests.Add($gRequestTmp)
        }
        if ($TeamsDevice.healthStatus -in $StatusType -or $SoftwareVersion) {
            $devicesWithUpdatePending++
            $gRequestTmp = New-Object -TypeName PSObject -Property @{
                id     = $TeamsDevice.id + "-health"
                method = "GET"
                url    = "/teamwork/devices/" + $TeamsDevice.id + "/health"
            }
            [void]$graphRequests.Add($gRequestTmp)
        }
        $gRequestTmp = New-Object -TypeName PSObject -Property @{
            id     = $TeamsDevice.id + "-operations"
            method = "GET"
            url    = "/teamwork/devices/" + $TeamsDevice.id + "/operations"
        }
        [void]$graphRequests.Add($gRequestTmp) 
    }
    if ($graphRequests.Count -gt 0) {
        $graphResponseExtra = Invoke-UcGraphRequest -Requests $graphRequests -Beta -Activity "Update-UcTeamsDevices, getting device health info" -IncludeBody
    }

    #In case we detect more than 5 devices with updates pending we will request confirmation that we can continue.
    if (($devicesWithUpdatePending -ge 5) -and !$ReportOnly) {
        if ($ConfirmPreference) {
            $title = 'Confirm'
            $question = "There are " + $devicesWithUpdatePending + " Teams Devices pending update. Are you sure that you want to continue?"
            $choices = '&Yes', '&No'
            $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
        }
        else {
            $decision = 0
        }
        if ($decision -ne 0) {
            return
        }
    }
    $graphRequests = [System.Collections.ArrayList]::new()
    foreach ($TeamsDevice in $TeamsDeviceList) {
        if ($TeamsDevice.healthStatus -in $StatusType -or $SoftwareVersion) {
            $TeamsDeviceHealth = ($graphResponseExtra | Where-Object { $_.id -eq ($TeamsDevice.id + "-health") }).body
                
            #Valid types are: adminAgent, operatingSystem, teamsClient, firmware, partnerAgent, companyPortal.
            #Currently we only consider Firmware and TeamsApp(teamsClient)
             
            #region Firmware
            if (($TeamsDeviceHealth.softwareUpdateHealth.firmwareSoftwareUpdateStatus.softwareFreshness.Equals("updateAvailable") -or $SoftwareVersion) -and ($UpdateType -in ("All", "Firmware"))) {
                if (!($ReportOnly)) {

                    $requestHeader = New-Object 'System.Collections.Generic.Dictionary[string, string]'
                    $requestHeader.Add("Content-Type", "application/json")
                    $requestBody = New-Object 'System.Collections.Generic.Dictionary[string, string]'
                    $requestBody.Add("softwareType", "firmware")
                        
                    if ($SoftwareVersion) {
                        $requestBody.Add("softwareVersion", $SoftwareVersion)
                    }
                    else {
                        $requestBody.Add("softwareVersion", $TeamsDeviceHealth.softwareUpdateHealth.firmwareSoftwareUpdateStatus.availableVersion)
                    }

                    $gRequestTmp = New-Object -TypeName PSObject -Property @{
                        id      = $TeamsDevice.id + "-updateFirmware"
                        method  = "POST"
                        url     = "/teamwork/devices/" + $TeamsDevice.id + "/updateSoftware"
                        body    = $requestBody
                        headers = $requestHeader
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                }
            }
            #endregion
                
            #region TeamsApp
            if (($TeamsDeviceHealth.softwareUpdateHealth.teamsClientSoftwareUpdateStatus.softwareFreshness.Equals("updateAvailable") -or $SoftwareVersion) -and ($UpdateType -in ("All", "TeamsClient"))) {
                if (!($ReportOnly)) {
                    $requestHeader = New-Object 'System.Collections.Generic.Dictionary[string, string]'
                    $requestHeader.Add("Content-Type", "application/json")

                    $requestBody = New-Object 'System.Collections.Generic.Dictionary[string, string]'
                    $requestBody.Add("softwareType", "teamsClient")
                    if ($SoftwareVersion) {
                        $requestBody.Add("softwareVersion", $SoftwareVersion)
                    }
                    else {
                        $requestBody.Add("softwareVersion", $TeamsDeviceHealth.softwareUpdateHealth.teamsClientSoftwareUpdateStatus.availableVersion)
                    }
                    $gRequestTmp = New-Object -TypeName PSObject -Property @{
                        id      = $TeamsDevice.id + "-updateTeamsClient"
                        method  = "POST"
                        url     = "/teamwork/devices/" + $TeamsDevice.id + "/updateSoftware"
                        body    = $requestBody
                        headers = $requestHeader
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                }
            }
            #endregion
        }
    }
    if ($graphRequests.Count -gt 0) {
        $updateGraphResponse = Invoke-UcGraphRequest-Requests $graphRequests Beta -Activity "Update-UcTeamsDevices, sending update commands" -IncludeBody
    }
    foreach ($TeamsDevice in $TeamsDeviceList) {
        if ($TeamsDevice.healthStatus -in $StatusType -or $SoftwareVersion) {
            $TeamsDeviceHealth = ($graphResponseExtra | Where-Object { $_.id -eq ($TeamsDevice.id + "-health") }).body
            if ($ReportOnly) {
                $UpdateStatus = "Report Only:"
                $pendingUpdate = $false
                if ($TeamsDeviceHealth.softwareUpdateHealth.firmwareSoftwareUpdateStatus.softwareFreshness.Equals("updateAvailable") -and ($UpdateType -in ("All", "Firmware"))) {
                    $UpdateStatus += " Firmware Update Pending;"
                    $pendingUpdate = $true
                }
                if ($TeamsDeviceHealth.softwareUpdateHealth.teamsClientSoftwareUpdateStatus.softwareFreshness.Equals("updateAvailable") -and ($UpdateType -in ("All", "TeamsClient"))) {
                    $UpdateStatus += " Teams App Update Pending;"
                    $pendingUpdate = $true
                }
                if (!$pendingUpdate) {
                    $UpdateStatus = "Report Only: No firmware or Teams App updates pending."
                }
            }
            else {
                $tmpUpdateStatus = ($updateGraphResponse | Where-Object { $_.id -eq ($TeamsDevice.id + "-updateFirmware") })
                $tmpUpdateRequest = $graphRequests | Where-Object { $_.id -eq ($TeamsDevice.id + "-updateFirmware") }
                if ($tmpUpdateStatus.status -eq 202) {
                    $UpdateStatus = "Firmware update request to version " + $tmpUpdateRequest.body.softwareVersion + " was queued."
                }
                elseif ($tmpUpdateStatus.status -eq 404) {
                    $UpdateStatus = "Unknown/Invalid firmware version: " + $tmpUpdateRequest.body.softwareVersion
                }
                elseif ($tmpUpdateStatus.status -eq 409) {
                    $UpdateStatus = "There is a firmware update pending, please check the update status."
                }
                $tmpUpdateStatus = ($updateGraphResponse | Where-Object { $_.id -eq ($TeamsDevice.id + "-updateTeamsClient") })
                $tmpUpdateRequest = $graphRequests | Where-Object { $_.id -eq ($TeamsDevice.id + "-updateTeamsClient") }
                if ($tmpUpdateStatus.status -eq 202) {
                    $UpdateStatus = "Teams App update request to version " + $tmpUpdateRequest.body.softwareVersion + " was queued."
                }
                elseif ($tmpUpdateStatus.status -eq 404) {
                    $UpdateStatus = "Unknown/Invalid Teams App version: " + $tmpUpdateRequest.body.softwareVersion
                }
                elseif ($tmpUpdateStatus.status -eq 409) {
                    $UpdateStatus = "There is a Teams App update pending, please check the update status."
                }
            }
            $userUPN = ($graphResponseExtra | Where-Object { $_.id -eq $TeamsDevice.currentuser.id }).body.userPrincipalName
            $TeamsDeviceOperations = ($graphResponseExtra | Where-Object { $_.id -eq ($TeamsDevice.id + "-operations") }).body.value
            $LastUpdateStatus = ""
            $LastUpdateInitiatedBy = ""
            $LastUpdateModifiedDate = ""

            #In this case we only need the last time we tried to udpdate.
            foreach ($TeamsDeviceOperation in $TeamsDeviceOperations) {
                if ($TeamsDeviceOperation.operationType -eq 'softwareUpdate') {
                    $LastUpdateStatus = $TeamsDeviceOperation.status
                    $LastUpdateInitiatedBy = $TeamsDeviceOperation.createdBy.user.displayName
                    $LastUpdateModifiedDate = $TeamsDeviceOperation.lastActionDateTime
                    break;
                }
            }
            $TDObj = [PSCustomObject][Ordered]@{
                TACDeviceID                     = $TeamsDevice.id
                DeviceType                      = Convert-UcTeamsDeviceType $TeamsDevice.deviceType
                Manufacturer                    = $TeamsDevice.hardwaredetail.manufacturer
                Model                           = $TeamsDevice.hardwaredetail.model
                UserDisplayName                 = $TeamsDevice.currentUser.displayName
                UserUPN                         = $userUPN
                HealthStatus                    = $TeamsDevice.healthStatus
                TeamsAdminAgentCurrentVersion   = $TeamsDeviceHealth.softwareUpdateHealth.adminAgentSoftwareUpdateStatus.currentVersion
                TeamsAdminAgentAvailableVersion = $TeamsDeviceHealth.softwareUpdateHealth.adminAgentSoftwareUpdateStatus.availableVersion
                FirmwareCurrentVersion          = $TeamsDeviceHealth.softwareUpdateHealth.firmwareSoftwareUpdateStatus.currentVersion
                FirmwareAvailableVersion        = $TeamsDeviceHealth.softwareUpdateHealth.firmwareSoftwareUpdateStatus.availableVersion
                CompanyPortalCurrentVersion     = $TeamsDeviceHealth.softwareUpdateHealth.companyPortalSoftwareUpdateStatus.currentVersion
                CompanyPortalAvailableVersion   = $TeamsDeviceHealth.softwareUpdateHealth.companyPortalSoftwareUpdateStatus.availableVersion
                OEMAgentAppCurrentVersion       = $TeamsDeviceHealth.softwareUpdateHealth.partnerAgentSoftwareUpdateStatus.currentVersion
                OEMAgentAppAvailableVersion     = $TeamsDeviceHealth.softwareUpdateHealth.partnerAgentSoftwareUpdateStatus.availableVersion
                TeamsAppCurrentVersion          = $TeamsDeviceHealth.softwareUpdateHealth.teamsClientSoftwareUpdateStatus.currentVersion
                TeamsAppAvailableVersion        = $TeamsDeviceHealth.softwareUpdateHealth.teamsClientSoftwareUpdateStatus.availableVersion
                PreviousUpdateStatus            = $LastUpdateStatus
                PreviousUpdateInitiatedBy       = $LastUpdateInitiatedBy
                PreviousUpdateModifiedDate      = $LastUpdateModifiedDate
                UpdateStatus                    = $UpdateStatus
            }
            [void]$outTeamsDevices.Add($TDObj)
        }
    }

    if ($outTeamsDevices.Count -eq 1) {
        return $outTeamsDevices
    }
    elseif ($outTeamsDevices.Count -gt 1) {
        $outTeamsDevices | Sort-Object DeviceType, Manufacturer, Model | Export-Csv -path $OutputFullPath -NoTypeInformation
        Write-Host ("Results available in: " + $OutputFullPath) -ForegroundColor Cyan
    }
    else {
        Write-Host ("No Teams Device(s) found that have pending update.") -ForegroundColor Cyan
    }   
}

