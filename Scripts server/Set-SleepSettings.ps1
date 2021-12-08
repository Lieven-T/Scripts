if ($env:COMPUTERNAME -match "LAPTOP|-LT|PC-TROLI2") {
    powercfg -H OFF
    powercfg -x -standby-timeout-ac 30
    powercfg -x -standby-timeout-dc 30
    powercfg -x -monitor-timeout-ac 30
    powercfg -x -monitor-timeout-dc 30
    powercfg -x -hibernate-timeout-ac 90
    powercfg -x -hibernate-timeout-dc 90
} else {
    powercfg -H OFF
    powercfg -x -standby-timeout-ac 0
    powercfg -x -monitor-timeout-ac 30
    powercfg -x -monitor-timeout-dc 30
}

If (Get-Command Get-IntelNetAdapterSetting -ErrorAction Ignore) {
	Get-IntelNetAdapterSetting -DisplayName "Wake on Pattern Match" | Set-IntelNetAdapterSetting -DisplayValue "Disabled"
	Get-IntelNetAdapterSetting -DisplayName "Wake on Magic Packet" | Set-IntelNetAdapterSetting -DisplayValue "Enabled"
} else {
	Set-NetAdapterPowerManagement -WakeOnMagicPacket "Enabled" -Name "Ethernet"
	Set-NetAdapterPowerManagement -WakeOnPattern "Disabled" -Name "Ethernet"
}