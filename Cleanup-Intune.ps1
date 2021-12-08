$byod = Get-AzureADGroup -SearchString oud-BYOD-laptops | Get-AzureADGroupMember
$byodDeviceId = $byod | select -ExpandProperty DeviceId
$byodObjectId = $byod | select -ExpandProperty ObjectId
$byodName = $byod | select -ExpandProperty DisplayName

$alldevices = Get-AzureADGroup -SearchString intune-campus-all-devices | Get-AzureADGroupMember
$allDeviceId = $alldevices | select -ExpandProperty DeviceId
$allObjectId = $alldevices | select -ExpandProperty ObjectId
$allName = $alldevices | select -ExpandProperty DisplayName

$joinDeviceId = $allDeviceId + $byodDeviceId
$joinObjectId = $allObjectId + $byodObjectId
$joinName = $allName + $byodName


$data = Get-IntuneManagedDevice
$data.value | ? deviceName -notin $joinName | % { Write-Host $_.deviceName; Remove-IntuneManagedDevice -managedDeviceId $_.id }
 
$data = Get-IntuneManagedDevice
$data.value | ? deviceName -notin $joinName | % { Write-Host $_.deviceName; Remove-IntuneManagedDevice -managedDeviceId $_.id }

$data = Get-IntuneManagedDevice
$data.value | ? deviceName -notin $joinName | % { Write-Host $_.deviceName; Remove-IntuneManagedDevice -managedDeviceId $_.id }

$fullList = Get-AzureADDevice -All $true

$intunedevices = $data | select -ExpandProperty devicename