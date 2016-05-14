#settings
$ip_eth0 = "10.0.0.5"
$alias_eth0 = "Ethernet"

$ip_wifi = "10.0.0.6"
$alias_wifi = "Wi-Fi"

$dns = "8.8.8.8","8.8.4.4"

$gateway = "10.0.0.1"

#remove previous IP settings (will only run if they exist)
#uses -Confirm:$false to automatically answer "Y" to confirmation prompts
Remove-NetIPAddress -InterfaceAlias $alias_eth0 -IPAddress $ip_eth0 -confirm:$false
Remove-NetIPAddress -InterfaceAlias $alias_wifi -IPAddress $ip_wifi -confirm:$false
Remove-NetRoute -DestinationPrefix 0.0.0.0/0 -confirm:$false

#set eth0 ip/dns
New-NetIPAddress -InterfaceAlias $alias_eth0 -AddressFamily IPv4 -IPAddress $ip_eth0 -PrefixLength 26 
Set-DnsClientServerAddress -InterfaceAlias $alias_eth0 -ServerAddresses $dns

#set wifi ip/dns
New-NetIPAddress -InterfaceAlias $alias_wifi -AddressFamily IPv4 -IPAddress $ip_wifi -PrefixLength 26 
Set-DnsClientServerAddress -InterfaceAlias $alias_wifi -ServerAddresses $dns

#recreate Default Gateway settings for redundant default routes
#note: must set routing table directly via New-NetRoute in order to achieve multiple redundant default routes
#note: metric 256 seems to be Window's default
New-NetRoute -DestinationPrefix 0.0.0.0/0 -InterfaceAlias $alias_wifi -NextHop $gateway -RouteMetric 256
New-NetRoute -DestinationPrefix 0.0.0.0/0 -InterfaceAlias $alias_eth0 -NextHop $gateway -RouteMetric 256

# DEBUG
Write-Host "Settings changed to:"
Get-NetIPAddress -InterfaceAlias $alias_eth0,$alias_wifi -AddressFamily IPv4  | Format-Table ifIndex, InterfaceAlias, IPAddress, PrefixLength, AddressState
Get-DnsClientServerAddress -InterfaceAlias $alias_eth0,$alias_wifi -AddressFamily IPv4 | Format-Table InterfaceIndex, InterfaceAlias, ServerAddresses

pause
