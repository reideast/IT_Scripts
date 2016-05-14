$alias_eth0 = "Ethernet"
$alias_wifi = "Wi-Fi"

Write-Host "Settings changed to:"
Get-NetIPAddress -InterfaceAlias $alias_eth0,$alias_wifi -AddressFamily IPv4  | Format-Table ifIndex, InterfaceAlias, IPAddress, PrefixLength, AddressState
Get-DnsClientServerAddress -InterfaceAlias $alias_eth0,$alias_wifi -AddressFamily IPv4 | Format-Table InterfaceIndex, InterfaceAlias, ServerAddresses

pause
