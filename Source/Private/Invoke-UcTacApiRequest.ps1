function Invoke-UcTacApiRequest {
    param(
        [string]$Uri
    )
    $isTokenValid = $false
    
    #Check if the current token is still valid. We use the Silent switch, if we have a valid token we can use it for the current request, if not an exception will be raise and we will prompt the user to authenticate.
    try {
        $AuthToken = Get-MsalToken -ClientId '12128f48-ec9e-42f0-b203-ea49fb6af367' -Scope 'https://devicemgmt.teams.microsoft.com/.default' -RedirectUri "https://teamscmdlet.microsoft.com" -Silent
        $isTokenValid = $true
    }
    catch {}

    #For now we will request a new token if the previous is expired.
    if (!$isTokenValid) {
        Write-Warning "Requesting a new authentication token."
        $AuthToken = Get-MsalToken -ClientId '12128f48-ec9e-42f0-b203-ea49fb6af367' -Scope 'https://devicemgmt.teams.microsoft.com/.default' -RedirectUri "https://teamscmdlet.microsoft.com"
        $isTokenValid = $true
    }

    if ($isTokenValid) {
        try {
            $header = @{Authorization = "Bearer " + $AuthToken.AccessToken }
            $TacApiResponse = Invoke-RestMethod -Uri $Uri -Method GET -Headers $header -ContentType "application/json"
        }
        catch {
            #TODO: Improve the error handling
            Write-Warning ("Failed to get a response from Graph API with the following message: " + [Environment]::NewLine + $_.Exception.Message)
            return
        }
    }
    return $TacApiResponse
}