# Enable firewall for all profiles
Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled True

# Disable firewall for all profiles
Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled False

Get-NetFirewallRule | Select-Object DisplayName, Enabled, Direction, Action

Get-NetFirewallProfile | Select-Object Name, Enabled

#<$FPD = Get-NetFirewallProfile -Name Domain |
# Select-Object -ExpandProperty Enabled
#  ForEach-Object {if ($_) {'Active'} else {'Disabled'} }
#  write-host $FPD

 $FPD if ((Get-NetFirewallProfile -Name Domain).enabled) {'Active'} else {'Disabled'}

Get-NetConnectionProfile | Select-Object InterfaceAlias, NetworkCategory

Get-NetFirewallProfile | Select-Object Name, Enabled,
    @{Name='Active';Expression={($_.Name -eq (Get-NetConnectionProfile).NetworkCategory)}}



New-NetFirewallRule -DisplayName "Block SMTP Outbound" `
    -Direction Outbound `
    -Protocol TCP `
    -LocalPort 25 `
    -Action Block
