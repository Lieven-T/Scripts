if (-not (Get-InstalledModule windowsautopilotintune)) {
    Install-PackageProvider -Name "NuGet" -Force
    Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
    Install-Module windowsautopilotintune
}

Connect-MSGraph

$AutopilotDevices = Get-AutopilotDevice
$SignpostDevices = Import-Excel -Path .\devicelist.xlsx
$SignpostDevices | % {
    $SignpostDevice = $_
    Write-Host "Instellen Autopilotnaam voor $($SignpostDevice.Label)"
    $AutopilotDevices | ? serialNumber -eq $SignpostDevice.Serienummer | % {
        Write-Host "    Gevonden device: $($_.id) in levering $($SignpostDevice.Leveringsnummer)"
        Set-AutopilotDevice -displayName $SignpostDevice.Label -id $_.id -groupTag $SignpostDevice.Leveringsnummer
    }
}


<#(Invoke-RESTMethod -Uri "http://orc-inventory/query" -UseBasicParsing).computers | % { 
    $InventoryDevice = $_
    $AutopilotDevices | ? serialNumber -eq $InventoryDevice.serialNumber | % {
        Write-Host "Instellen Autopilotnaam voor $($InventoryDevice.hostname)"
        Write-Host "    Gevonden device: $($_.id)"
        Set-AutopilotDevice -displayName $InventoryDevice.hostname -id $_.id
    }
}#>
