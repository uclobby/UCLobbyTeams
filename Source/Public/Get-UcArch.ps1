function Get-UcArch {
    param(
        [string]$FilePath
    )
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