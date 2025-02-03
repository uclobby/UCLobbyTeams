function Get-UcEntraObjectsOwnedByUser {
    <#
        .SYNOPSIS
        Returns all Entra objects associated with a user.

        .DESCRIPTION
        This function returns the list of Entra Objects associated with the specified user.

        Contributors: Jimmy Vincent, David Paulino

        Requirements:   Microsoft Graph Authentication PowerShell Module (Install-Module Microsoft.Graph.Authentication)
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
    if (Test-UcMgGraphConnection -Scopes "User.Read.All" -AltScopes "Directory.Read.All") {
        #2025-01-31: Only need to check this once per PowerShell session
        if (!($global:UCLobbyTeamsModuleCheck)) {
            Test-UcPowerShellModule -ModuleName UcLobbyTeams | Out-Null
            $global:UCLobbyTeamsModuleCheck = $true
        }
        $output = [System.Collections.ArrayList]::new()
        $graphRequests = [System.Collections.ArrayList]::new()
        $gRequestTmp = New-Object -TypeName PSObject -Property @{
            id     = "UserObjects"
            method = "GET"
            url    = "/users/$User/ownedObjects?`$select=id,displayName,visibility,createdDateTime,creationOptions,groupTypes,securityEnabled,mailEnabled"
        }
        [void]$graphRequests.Add($gRequestTmp)
        $BatchResponse = Invoke-UcMgGraphBatch -Requests $graphRequests -IncludeBody 
        
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
                $UserObject = New-Object -TypeName PSObject -Property @{
                    User            = $User
                    ObjectID        = $OwnedObject.id
                    DisplayName     = $OwnedObject.displayName
                    Type            = $tmpType
                    CreatedDateTime = $OwnedObject.createdDateTime
                    Visibility      = $OwnedObject.visibility
                }
                $UserObject.PSObject.TypeNames.Insert(0, 'EntraObjectsOwnedByUser')
                [void]$output.Add($UserObject)
            }
        }
        return $output | Sort-Object Type, DisplayName
    }
}