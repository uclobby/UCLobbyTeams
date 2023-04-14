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

Function Get-UcTeamsVersionBatch {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$InputCSV,
        [string]$OutputPath,
        [switch]$ExportCSV,
        [System.Management.Automation.PSCredential]$Credential
    )
    if (Test-Path $InputCSV) {
        try{
            $Computers = Import-Csv -Path $InputCSV
        } catch {
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
            $tmpTV = Get-UcTeamsVersion -Computer $computer.Computer -Credential $cred
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
    } else {
        Write-Host ("Error: File not found " + $InputCSV) -ForegroundColor Red
    }
}