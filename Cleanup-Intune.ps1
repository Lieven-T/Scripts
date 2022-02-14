$byod = Get-AzureADGroup -SearchString oud-BYOD-laptops | Get-AzureADGroupMember
$byodName = $byod | select -ExpandProperty DisplayName

$alldevices = Get-AzureADGroup -SearchString intune-campus-all-devices | Get-AzureADGroupMember
$allName = $alldevices | select -ExpandProperty DisplayName

$joinName = $allName + $byodName


$data = Get-IntuneManagedDevice
$data.value | ? deviceName -notin $joinName | % { Write-Host $_.deviceName; Remove-IntuneManagedDevice -managedDeviceId $_.id }
 
$data = Get-IntuneManagedDevice
$data.value | ? deviceName -notin $joinName | % { Write-Host $_.deviceName; Remove-IntuneManagedDevice -managedDeviceId $_.id }

$data = Get-IntuneManagedDevice
$data.value | ? deviceName -notin $joinName | % { Write-Host $_.deviceName; Remove-IntuneManagedDevice -managedDeviceId $_.id }
