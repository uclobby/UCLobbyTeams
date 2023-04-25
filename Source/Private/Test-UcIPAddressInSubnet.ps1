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

Function Test-UcIPaddressInSubnet {

    param(
        [string]$IPAddress,
        [string]$Subnet
    )

    $regExIPAddressSubnet = "^((25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9])\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[0-9]))\/(3[0-2]|[1-2]{1}[0-9]{1}|[1-9])$"

    try {
        $Subnet -match $regExIPAddressSubnet | Out-Null

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