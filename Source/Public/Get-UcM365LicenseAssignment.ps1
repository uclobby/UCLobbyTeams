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
        [switch]$UseFriendlyNames,
        [switch]$SkipServicePlan,
        [string]$OutputPath
    )

    if ((Test-UcMgGraphConnection -Scopes "Directory.Read.All" -AltScopes ("User.Read.All","Group.Read.All","Organization.Read.All"))) {
        Test-UcModuleUpdateAvailable -ModuleName UcLobbyTeams
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
            $SKUnSPFilePath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads", "Product names and service plan identifiers forlicensing.csv")
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
            url    = "/subscribedSkus?`$select=skuID,skuPartNumber,servicePlans,appliesTo"
        }
        [void]$graphRequests.Add($gRequestTmp)
        $TenantSKUs = (Invoke-UcMgGraphBatch -Requests $graphRequests -MgProfile beta -Activity "Get-UcM365LicenseAssignment, Step 1: Getting Tenant License information").value

        $graphHeader = New-Object 'System.Collections.Generic.Dictionary[string, string]'
        $graphHeader.Add("ConsistencyLevel", "eventual")
        
        $graphRequests = [System.Collections.ArrayList]::new()
        $gRequestTmp = New-Object -TypeName PSObject -Property @{
            id     = "GroupsWithLicenses"
            method = "GET"
            headers= $graphHeader
            url    = "/groups?`$filter=assignedLicenses/`$count ne 0&`$count=true&`$select=id,displayName,assignedLicenses&`$top=999"
        }
        [void]$graphRequests.Add($gRequestTmp)
        $GroupsWithLicenses = (Invoke-UcMgGraphBatch -Requests $graphRequests -MgProfile beta -Activity "Get-UcM365LicenseAssignment, Step 2: Getting Group with licenses assigned").value

        $graphRequests = [System.Collections.ArrayList]::new()
        $gRequestTmp = New-Object -TypeName PSObject -Property @{
            id     = "UsersWithLicenses"
            method = "GET"
            headers= $graphHeader
            url    = "/users?`$filter=assignedLicenses/`$count ne 0&`$count=true&`$select=userPrincipalName,licenseAssignmentStates&`$top=999"
        }
        [void]$graphRequests.Add($gRequestTmp)
        $UsersWithLicenses = (Invoke-UcMgGraphBatch -Requests $graphRequests -MgProfile beta -Activity "Get-UcM365LicenseAssignment, Step 3: Getting User License information").value
        $userProcessed = 1
        $totalUsers = $UsersWithLicenses.count
        $outUserLicenseAssignment = [System.Collections.ArrayList]::new()
        
        foreach($UserWithLicense in $UsersWithLicenses){
            #Only update evey 1000 users processed
            if(($userProcessed%1000 -eq 0) -or ($userProcessed -eq $totalUsers)){
                Write-Progress -Activity "Get-UcM365LicenseAssignment, Step 4: Reading users assigned licenses/service plans" -Status "$userProcessed of $totalUsers"
            }
            foreach ($licenseState in $UserWithLicense.licenseAssignmentStates) {
                $licenseAssignment = "Direct"
                $licenseAssignmentGroup = ""
    
                if (!([string]::IsNullOrEmpty($licenseState.assignedByGroup))) {
                    $licenseAssignment = "Inherited"
                    $licenseAssignmentGroup = ($GroupsWithLicenses | Where-Object -Property "id" -EQ -Value $licenseState.assignedByGroup).displayName
                }
                
                $SKU = $TenantSKUs | Where-Object { $_.skuID -eq $licenseState.skuId }
                if(!($SkipServicePlan)){
                    $tmpUserServicePlan = [System.Collections.ArrayList]::new()
                    foreach($servicePlan in $SKU.servicePlans){
                        if($servicePlan.appliesTo -eq "User"){
                            if($servicePlan.servicePlanId -notin $licenseState.disabledPlans){
                                $ServicePlanStatus = "On"
                            } else {
                                $ServicePlanStatus = "Off"
                            }
                            $ULObj = New-Object -TypeName PSObject -Property @{
                                ServicePlanId = $servicePlan.servicePlanId
                                ServicePlanName = $servicePlan.servicePlanName
                                ServicePlanStatus = $ServicePlanStatus 
                            }
                            [void]$tmpUserServicePlan.Add($ULObj)
                        }
                    }
                } else {
                    $tmpUserServicePlan = "Service Plans skipped"
                }

                if ($SKU) {
                    $tmpUserServicePlan = $tmpUserServicePlan | Sort-Object ServicePlanName
                    $UTObj = New-Object -TypeName PSObject -Property @{
                        UserUPN                     = $UserWithLicense.userPrincipalName
                        License                     = $SKU.skuPartNumber
                        LicenseID                   = $licenseState.skuId
                        LicenseAssignment           = $licenseAssignment
                        LicenseAssignmentGroup      = $licenseAssignmentGroup
                        LicenseAssignmentGroupID    = $licenseState.assignedByGroup
                        UserServicePlans            = $tmpUserServicePlan
                    }
                    $UTObj.PSObject.TypeNAmes.Insert(0, 'UserLicenseAssignment')
                    [void]$outUserLicenseAssignment.Add($UTObj)
                }
            }
            $userProcessed++
        }

        if ($outUserLicenseAssignment) {
            if($SkipServicePlan){
                $outUserLicenseAssignment | Sort-Object UserUPN, License, LicenseAssignment | Select-Object UserUPN, License, LicenseAssignment, LicenseAssignmentGroup | Export-Csv -Path $OutputFilePath -NoTypeInformation
            } else {
                $headerstring = "UserPrincipalName,LicenseAssigned,LicenseAssignment,LicenseAssignmentGroup"
                $lastSKU = ""

                $LAProcessed = 1
                $totalLicenseAssignments = $outUserLicenseAssignment.Count

                $outUserLicenseAssignment  = $outUserLicenseAssignment | Sort-Object License,UserUPN,LicenseAssignment
                foreach($userEntry in $outUserLicenseAssignment){
                    #Only update activity every 1000 entries
                    if(($LAProcessed%1000 -eq 0) -or ($LAProcessed -eq $totalLicenseAssignments)){
                        Write-Progress -Activity "Get-UcM365LicenseAssignment, Step 5: Writing license assignments" -Status "$LAProcessed of $totalLicenseAssignments"
                    }
                    $row = ""
                    if($lastSKU -ne $userEntry.LicenseID){
                        $lastSKU = $userEntry.LicenseID
                        $row += [Environment]::NewLine + $headerstring
                        if($UseFriendlyNames){
                            $LicenseDisplayName = ($SKUnSP | Where-Object { $_.GUID -eq $userEntry.LicenseID }|Sort-Object Product_Display_Name -Unique).Product_Display_Name
                        } else {
                            $LicenseDisplayName = $userEntry.License
                        }
                        foreach($plan in $userEntry.UserServicePlans){
                            if($UseFriendlyNames){
                                $servicePlanName = ($SKUnSP | Where-Object { $_.Service_Plan_Id -eq $plan.ServicePlanId  } | Sort-Object Service_Plans_Included_Friendly_Names -Unique).Service_Plans_Included_Friendly_Names
                                if([string]::IsNullOrEmpty($servicePlanName)){
                                    $servicePlanName = $plan.servicePlanName    
                                }
                            } else {
                                $servicePlanName = $plan.ServicePlanName
                            }
                            $row += "," + $servicePlanName 
                        }
                        $row += [Environment]::NewLine
                    } 
                    $row += $userEntry.UserUPN + "," + $LicenseDisplayName + "," + $userEntry.LicenseAssignment + "," + $userEntry.LicenseAssignmentGroup
                    foreach($plan in $userEntry.UserServicePlans){
                        $row += "," + $plan.ServicePlanStatus
                    }
                    Out-File -FilePath $OutputFilePath -InputObject $row -Encoding UTF8 -append
                    $LAProcessed++    
                }
            }
        }
        Write-Host ("Results available in " + $OutputFilePath) -ForegroundColor Cyan
    }
}