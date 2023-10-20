<#
.SYNOPSIS
Generate a report of the User assigned licenses either direct or assigned by group (Inherited)

.DESCRIPTION
This script will get a report of all Service Plans assigned to users and how the license is assigned to the user (Direct, Inherited)

Contributors: David Paulino, Freydem Fernandez Lopez, Gal Naor

Requirements:   Microsoft Graph PowerShell Module (Install-Module Microsoft.Graph)
                Microsoft Graph Scopes:
                        "Directory.Read.All"
.PARAMETER UseFriendlyNames
When present will download a csv file containing the License/ServicePlans friendly names

Product names and service plan identifiers for licensing
https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference

.PARAMETER SkipServicePlan
When present will just check the licenses and not the service plans assigned to the user.

.PARAMETER OutputPath
Allows to specify the path where we want to save the results. By default, it will save on current user Download.

.EXAMPLE 
PS> Get-UcM365LicenseAssignment

.EXAMPLE 
PS> Get-UcM365LicenseAssignment -UseFriendlyNames

#>

Function Get-UcM365LicenseAssignment {
    Param(
        [string]$SKU,    
        [switch]$UseFriendlyNames,
        [switch]$SkipServicePlan,
        [string]$OutputPath
    )

    if ((Test-UcMgGraphConnection -Scopes "Directory.Read.All" -AltScopes ("User.Read.All","Organization.Read.All"))) {
        Test-UcModuleUpdateAvailable -ModuleName UcLobbyTeams
        $startTime=Get-Date;
        $outFile = "M365LicenseAssigment_" + (Get-Date).ToString('yyyyMMdd-HHmmss') + ".csv"
        #Verify if the Output Path exists
        if ($OutputPath) {
            if (!(Test-Path $OutputPath -PathType Container)) {
                Write-Host ("Error: Invalid folder " + $OutputPath) -ForegroundColor Red
                return
            }
            $OutputFilePath = [System.IO.Path]::Combine($OutputPath, $outFile)
        }
        else {                
            $OutputFilePath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads", $outFile)
        }
        
        if($UseFriendlyNames){
            #20231019 - Change: OutputPath will be for both report and Product names and service plan identifiers for licensing.csv
            $SKUnSPFilePath = [System.IO.Path]::Combine($OutputPath, "Product names and service plan identifiers for licensing.csv")
            if (!(Test-Path -Path $SKUnSPFilePath)) {
                try {
                    Write-Warning "M365 Product Names and Service Plans file not found, attempting to download it."
                    Invoke-WebRequest -Uri "https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv" -OutFile $SKUnSPFilePath
                }
                catch {
                    Write-Warning "Could not download M365 Product Names and Service Plans."
                }
            }
            try {
                $SKUnSP = import-CSV -Path $SKUnSPFilePath
            }
            catch {
                Write-Warning "Could not import Service Plan ID file."
                $UseFriendlyNames = $false
            }
        }
        $graphRequests = [System.Collections.ArrayList]::new()
        $gRequestTmp = New-Object -TypeName PSObject -Property @{
            id     = "TenantSKUs"
            method = "GET"
            url    = "/subscribedSkus?`$select=skuID,skuPartNumber,servicePlans,appliesTo,consumedUnits"
        }
        [void]$graphRequests.Add($gRequestTmp)
        $TenantSKUs = (Invoke-UcMgGraphBatch -Requests $graphRequests -MgProfile beta -Activity "Get-UcM365LicenseAssignment, Step 1: Getting Tenant License information").value

        #region 20232019 - Adding filter to SKU
        if($SKU){
            if ($UseFriendlyNames){
                $SKUGUID = ($SKUnSP | Where-Object { $_.String_Id -eq $SKU -or $_.Product_Display_Name -eq $SKU} | Sort-Object GUID -Unique ).GUID
                $TenantSKUs = $TenantSKUs | Where-Object {$_.skuId -eq $SKUGUID}
                if($TenantSKUs.count -eq 0){
                    Write-Warning "Could not find `"$SKU`" (SKU Name/Part Number) subscription associated with the tenant."
                    return 
                }
            } else {
                $TenantSKUs = $TenantSKUs | Where-Object {$_.skuPartNumber -eq $SKU}
                if($TenantSKUs.count -eq 0){
                    Write-Warning "Could not find `"$SKU`" (SKU Part Number) subscription associated with the tenant."
                    return 
                }
            }
        } else {
            $TenantSKUs = $TenantSKUs | Where-Object -Property consumedUnits -GT -Value 0 | Sort-Object skuPartNumber
        }
        #endregion

        #region 20131019 - Getting all Service Plans for new matrix style report
        $allServicePlans = [System.Collections.ArrayList]::new()
        foreach($TenantSKU in $TenantSKUs){
            $tmpUserServicePlans = $TenantSKU.ServicePlans | Where-Object -Property appliesTo -EQ -Value "User" 
            foreach($ServicePlan in $tmpUserServicePlans){
                if(!($ServicePlan.ServicePlanId -in $allServicePlans.ServicePlanId)){
                    if($UseFriendlyNames){
                        $servicePlanName = ($SKUnSP | Where-Object { $_.Service_Plan_Id -eq $ServicePlan.ServicePlanId -and $_.GUID -eq $TenantSKU.skuID } | Sort-Object Service_Plans_Included_Friendly_Names -Unique).Service_Plans_Included_Friendly_Names
                        if([string]::IsNullOrEmpty($servicePlanName)){
                            $servicePlanName = $ServicePlan.servicePlanName    
                        }
                    } else {
                        $servicePlanName = $ServicePlan.ServicePlanName
                    }
                    $tmpSP = New-Object -TypeName PSObject -Property @{
                        servicePlanId       = $ServicePlan.ServicePlanId
                        servicePlanName     = $servicePlanName
                    }
                    [void]$allServicePlans.Add($tmpSP)
                }
            }
        }
        #Sorting service plans by name and creating the file header
        $allServicePlans = $allServicePlans | Sort-Object ServicePlanName
        $row = "UserPrincipalName,LicenseAssigned,LicenseAssignment,LicenseAssignmentGroup"
        if(!($SkipServicePlan)){
            foreach($ServicePlan in $allServicePlans){
                $row += "," +  $ServicePlan.servicePlanName
            }
        }
        $row += [Environment]::NewLine
        #endregion

        $GraphRequestHeader = New-Object 'System.Collections.Generic.Dictionary[string, string]'
        $GraphRequestHeader.Add("ConsistencyLevel", "eventual")
        $graphRequests = [System.Collections.ArrayList]::new()
        $gRequestTmp = New-Object -TypeName PSObject -Property @{
            id     = "GroupsWithLicenses"
            method = "GET"
            headers= $GraphRequestHeader
            url    = "/groups?`$filter=assignedLicenses/`$count ne 0&`$count=true&`$select=id,displayName,assignedLicenses&`$top=999"
        }
        [void]$graphRequests.Add($gRequestTmp)
        $GroupsWithLicenses = (Invoke-UcMgGraphBatch -Requests $graphRequests -MgProfile beta -Activity "Get-UcM365LicenseAssignment, Step 2: Getting Group with licenses assigned").value
        
        Write-Progress -Id 3 -Activity "Get-UcM365LicenseAssignment, Step 3: Reading users assigned licenses/service plans"
        foreach($TenantSKU in $TenantSKUs)
        {
            if($UseFriendlyNames){
                $LicenseDisplayName = ($SKUnSP | Where-Object { $_.GUID -eq $TenantSKU.skuID } | Sort-Object Product_Display_Name -Unique).Product_Display_Name
            } else {
                $LicenseDisplayName = $TenantSKU.skuPartNumber
            }
            $SKUUserServicePlans = $TenantSKU.servicePlans | Where-Object -Property appliesTo -EQ -Value "User" | Sort-Object servicePlanName
            $usersProcessed = 0       
            $GraphRequestURI = "https://graph.microsoft.com/v1.0/users?`$filter=assignedLicenses/any(u:u/skuId eq " + $TenantSKU.skuId + " )&`$select=userPrincipalName,licenseAssignmentStates&`$orderby=userPrincipalName&`$count=true&`$top=999"
            do{
                try{
					$UsersWithLicenses = Invoke-MgGraphRequest -Method Get -Uri $GraphRequestURI -Headers $GraphRequestHeader
					if (![string]::IsNullOrEmpty($UsersWithLicenses.'@odata.count')){
						$TotalUsers = $UsersWithLicenses.'@odata.count'
					}
					$GraphRequestURI = $UsersWithLicenses.'@odata.nextLink'
					foreach($UserWithLicense in $UsersWithLicenses.value){
						if(($usersProcessed%1000 -eq 0) -or ($usersProcessed -eq $TotalUsers)){
							Write-Progress -ParentId 3 -Activity "Checking license assignments for $LicenseDisplayName" -Status "$usersProcessed of $TotalUsers"
						}
						$tmpLicenseAssignmentStates = $UserWithLicense.licenseAssignmentStates | Where-Object -Property skuId -EQ -Value $TenantSKU.skuId | Sort-Object assignedByGroup
						foreach ($licenseState in $tmpLicenseAssignmentStates) {
							$licenseAssignment = "Direct"
							$licenseAssignmentGroup = ""
							if (!([string]::IsNullOrEmpty($licenseState.assignedByGroup))) {
								$licenseAssignment = "Inherited"
								$licenseAssignmentGroup = ($GroupsWithLicenses | Where-Object -Property "id" -EQ -Value $licenseState.assignedByGroup).displayName
								if([string]::IsNullOrEmpty($licenseAssignmentGroup)){
									$licenseAssignmentGroup = $licenseState.assignedByGroup
								}
							}
							$userServicePlans = ""
							if(!($SkipServicePlan)){
								foreach($ServicePlan in $allServicePlans){
									if($servicePlan.servicePlanId -in $SKUUserServicePlans.servicePlanId){
										if($servicePlan.servicePlanId -notin $licenseState.disabledPlans){
											$userServicePlans += ",On"
										} else {
											$userServicePlans += ",Off"
										}
									} else {
										$userServicePlans += ","
									}
								}
							}
							$row += $UserWithLicense.userPrincipalName + "," + $LicenseDisplayName + "," + $LicenseAssignment + "," + $LicenseAssignmentGroup + $userServicePlans
							Out-File -FilePath $OutputFilePath -InputObject $row -Encoding UTF8 -append
							$row = ""
						}
						$usersProcessed++
					}
				} catch {
						Write-Warning ("Failed to get Users with assigned SKU Id: " + $TenantSKU.skuID)
						$GraphRequestURI = ""
				}
				
            } while(![string]::IsNullOrEmpty($GraphRequestURI))
        }
        if($usersProcessed -gt 0){
            Write-Host ("Results available in " + $OutputFilePath) -ForegroundColor Cyan
            #region 20231019 - Change: Added execution time to the output.
            $endTime = Get-Date
            $totalSeconds= [math]::round(($endTime - $startTime).TotalSeconds,2)
            $totalTime = New-TimeSpan -Seconds $totalSeconds
            Write-Host "Execution time:" $totalTime.Hours "Hours" $totalTime.Minutes "Minutes" $totalTime.Seconds "Seconds" -ForegroundColor Green
            #endregion
        }
    }
}