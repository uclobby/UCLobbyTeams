function Invoke-UcMgGraphBatch {
    param(
        [object]$Requests,
        [ValidateSet("beta", "v1.0")]
        [string]$MgProfile,
        [string]$Activity,
        [switch]$IncludeBody
    )
    $GraphURI_BetaAPIBatch = "https://graph.microsoft.com/beta/`$batch"
    $GraphURI_ProdAPIBatch = "https://graph.microsoft.com/v1.0/`$batch"
    $outBatchResponses = [System.Collections.ArrayList]::new()
    $tmpGraphRequests = [System.Collections.ArrayList]::new()

    if ($MgProfile.Equals("beta")) {
        $GraphURI_Batch = $GraphURI_BetaAPIBatch
    }
    else {
        $GraphURI_Batch = $GraphURI_ProdAPIBatch
    }

    #If activity is null then we can use this to get the function that call this function.
    if (!($Activity)) {
        $Activity = [string]$(Get-PSCallStack)[1].FunctionName
    }

    $g = 1
    $batchCount = [int][Math]::Ceiling(($Requests.count / 20))
    foreach ($GraphRequest in $Requests) {
        Write-Progress -Activity $Activity -Status "Running batch $g of $batchCount"
        [void]$tmpGraphRequests.Add($GraphRequest) 
        if ($tmpGraphRequests.Count -ge 20) {
            $g++
            $grapRequestBody = ' { "requests":  ' + ($tmpGraphRequests  | ConvertTo-Json) + ' }' 
            $GraphResponses += (Invoke-MgGraphRequest -Method Post -Uri $GraphURI_Batch -Body $grapRequestBody).responses
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
        try{
        $GraphResponses += (Invoke-MgGraphRequest -Method Post -Uri $GraphURI_Batch -Body $grapRequestBody).responses
        } catch
        {
            Write-Warning "Error while getting the Graph Request."
        }
    }

    #In some cases we will need the complete graph response, in that case the calling function will have to process pending pages.
    $attempts = 1
    for ($j = 0; $j -lt $GraphResponses.length; $j++) {
        $ResponseCount = 0
        if($IncludeBody){
            $outBatchResponses += $GraphResponses[$j]
        } else {
            $outBatchResponses += $GraphResponses[$j].body
            if ($GraphResponses[$j].status -eq "200"){
                #Checking if there are more pages available    
                $GraphURI_NextPage = $GraphResponses[$j].body.'@odata.nextLink'
                $GraphTotalCount = $GraphResponses[$j].body.'@odata.count'
                $ResponseCount += $GraphResponses[$j].body.value.count
                while (![string]::IsNullOrEmpty($GraphURI_NextPage)) {
                    try{
                        $graphNextPageResponse = Invoke-MgGraphRequest -Method Get -Uri $GraphURI_NextPage
                        $outBatchResponses += $graphNextPageResponse
                        $GraphURI_NextPage = $graphNextPageResponse.'@odata.nextLink'
                        $ResponseCount += $graphNextPageResponse.value.count
                        Write-Progress -Activity $Activity -Status "$ResponseCount of $GraphTotalCount"
                    } catch {
                        Write-Warning "Failed to get the next batch page, retrying..."
                        $attempts--
                    }
                    if($attempts -eq 0){
                        Write-Warning "Could not get next batch page, skiping it."
                        break
                    }
                }
            } else {
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