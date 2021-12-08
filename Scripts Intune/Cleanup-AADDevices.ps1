Get-AzureADDevice -SearchString "User deleted for this device" -All $true | Remove-AzureADDevice

$SignpostDevices = Import-Excel -Path C:\temp\devicelist.xlsx
$SignpostLabels = $SignpostDevices | select -ExpandProperty Label
Get-AzureADDevice -SearchString "ORC1S" -All $true | ? { $_.DisplayName -notin $SignpostLabels } | Remove-AzureADDevice
Get-AutopilotDevice | ? { $_.DisplayName -match "ORC1S"  -and $_.DisplayName -notin $SignpostLabels } | Remove-AutopilotDevice