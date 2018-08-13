########################################################################
# Must run from elevated PS or ISE
# Easiest to drop into system32 folder
########################################################################

########################################################################
# DHCP Find
# Simple script, modify it as you need
########################################################################

### NOTES ###
# Requires Powershell 4.0 or higher, type "get-host" to verify
# Requires scripts enabled. Type "set-executionpolicy unrestricted" to fix

# Find DHCP Server
$DHCPServer = $(ipconfig /all | findstr /C:"DHCP Server")
$DHCPServer = $($DHCPServer -split ' ')[-1]
write-host "Connected to DHCP server $DHCPServer"

# Getting list of scopes
$scopes = $(netsh dhcp server $DHCPServer show scope)
$scopes = $scopes -match '^\s[0-9]' | %{$($_ -split ' ')[1] }

# Get Clients
write-host "Pulling all clients, this takes a few seconds`n (Sometimes Net Command Shell crashes, just click ok on that window)"
$clients = $null
$clients += $scopes | %{netsh dhcp server $DHCPServer scope $_ show clients}
$clients = $clients -replace '-', '' -match '^[0-9]' #Just to clean it up a bit

# Search
while (1) {
$search = Read-Host "`nWhat do you want to find? Examples: 10.10.10.10 or aabbccddee or ddee "
$clients -match "$search"
}
