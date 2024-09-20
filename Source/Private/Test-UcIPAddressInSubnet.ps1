function Test-UcIPaddressInSubnet {
    <#
        .SYNOPSIS
        Check if an IP address is part of an Subnet.

        .DESCRIPTION
        Returns true if the given IP address is part of the subnet, false for not or invalid ip address.
    
        Contributors: David Paulino

        .PARAMETER IPAddress
        IP Address that we want to confirm that belongs to a range.

        .PARAMETER Subnet
        Subnet in the IPaddress/SubnetMaskBits.

        .EXAMPLE 
        PS> Test-UcIPaddressInSubnet -IPAddress 192.168.0.1 -Subnet 192.168.0.0/24

    #>
    param(
        [Parameter(mandatory = $true)]    
        [string]$IPAddress,
        [Parameter(mandatory = $true)]
        [string]$Subnet
    )
    $regExIPAddressSubnet = "^((25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9])\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[0-9]))\/(3[0-2]|[1-2]{1}[0-9]{1}|[1-9])$"
    try {
        [void]($Subnet -match $regExIPAddressSubnet)
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