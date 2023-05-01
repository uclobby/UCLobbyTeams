Function Test-UcModuleUpdateAvailable {
    Param(
        [Parameter(Mandatory = $true)]    
        [string]$ModuleName
    )
    try { 
        #Get the current loaded version
        $tmpCurrentVersion = (get-module $ModuleName | Sort-Object Version -Descending)
        if ($tmpCurrentVersion){
            $currentVersion = $tmpCurrentVersion[0].Version.ToString()
        }
        
        #Get the lastest version available
        $availableVersion = (Find-Module -Name $ModuleName -Repository PSGallery -ErrorAction SilentlyContinue).Version
        #Get all installed versions
        $installedVersions = (get-module $ModuleName -ListAvailable).Version

        if ($currentVersion -ne $availableVersion ) {
            if ($availableVersion -in $installedVersions) {
                Write-Warning ("The lastest available version of $ModuleName module is installed, however version $currentVersion is imported." + [Environment]::NewLine + "Please make sure you import it with: Import-Module $ModuleName -RequiredVersion $availableVersion")
            }
            else {
                Write-Warning ("There is a new version available ($availableVersion), please update the module with: Update-Module $ModuleName")
            }
        }
    }
    catch {
    }
}