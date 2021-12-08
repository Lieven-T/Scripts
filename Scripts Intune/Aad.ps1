$allDevices = Get-AzureADDevice -All $true
$byod = get-azureadgroup -SearchString intune-byod | Get-AzureADGroupMember
$byodDevId = $byod | select -ExpandProperty DeviceId
$byodOud = get-azureadgroup -SearchString oud-BYOD-laptops | Get-AzureADGroupMember
$byodOudDevId = $byodOud | select -ExpandProperty DeviceId

$byodOnly = $byod | ? DeviceId -NotIn $byodOudDevId
$byodOudOnly = $byodOud | ? DeviceId -NotIn $byodDevId

$byodOud | ? ProfileType -eq registeredDevice
$byod | ? ProfileType -eq registeredDevice