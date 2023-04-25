Function Invoke-UcMgGraphBatch {
    Param(
        [object]$Requests,
        [ValidateSet("beta", "v1.0")]
        [string]$MgProfile,
        [string]$Activity
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
    $g = 1
    $batchCount = [int][Math]::Ceiling(($Requests.count / 20))
    foreach($GraphRequest in $Requests){
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
        $GraphResponses += (Invoke-MgGraphRequest -Method Post -Uri $GraphURI_Batch -Body $grapRequestBody).responses
    }

    for ($j = 0; $j -lt $GraphResponses.length; $j++) {
        $outBatchResponses += $GraphResponses[$j]
        #Checking if there are more pages available
        $GraphURI_NextPage = $GraphResponses[$j].body.'@odata.nextLink'
        while (![string]::IsNullOrEmpty($GraphURI_NextPage)) {
            $graphNextPageResponse = Invoke-MgGraphRequest -Method Get -Uri $GraphURI_NextPage
            $outBatchResponses += $graphNextPageResponse
            $GraphURI_NextPage = $graphNextPageResponse.'@odata.nextLink'
        }
    }
    return $outBatchResponses
}