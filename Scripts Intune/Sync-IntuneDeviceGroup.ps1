$Group = Get-AzureADGroup -SearchString "byod-laptops"
$CurrentMembers = Get-AzureADGroupMember -ObjectId $Group.ObjectId
$AllPc = Import-Excel C:\temp\devicelist.xlsx
$AllPc | % {
    $SerNum = $_.'Serial Number'
    $Pc = Get-IntuneManagedDevice -Filter "contains(serialNumber,'$SerNum')"
    Write-Host "PC met serienummer: $($Pc.serialNumber) en naam $($Pc.deviceName)"
    if ($Pc -and $Pc.serialNumber) {
        $ADDevice = Get-AzureADDevice -SearchString $Pc.deviceName
        Add-AzureADGroupMember -ObjectId $Group.ObjectId -RefObjectId $ADDevice.ObjectId
    }
}

$CurrentMembers | % {
    $HostName = $_.DisplayName
    $Pc = Get-IntuneManagedDevice -Filter "contains(deviceName,'$HostName')"
    Write-Host "PC met serienummer: $($Pc.serialNumber) en naam $($Pc.deviceName)"
    if ($Pc -and $Pc.serialNumber) {
        if ($Pc.serialNumber -notin $sernums) {
            Write-Host "    Niet aanwezig"
        }
    }
}
