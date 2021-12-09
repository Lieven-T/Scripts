Param
(
    [Parameter(Mandatory=$True,Position=1)]
    [ValidateScript({Test-Path $_})] 
    [string]$InputFile
)

if (-not (Get-InstalledModule windowsautopilotintune)) {
    Install-PackageProvider -Name "NuGet" -Force
    Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
    Install-Module windowsautopilotintune
}

Connect-MSGraph
$AutopilotDevices = Get-AutopilotDevice
Import-Excel -Path $InputFile | % {
    $SignpostDevice = $_
    Write-Host "Instellen Autopilotnaam voor $($SignpostDevice.Label)"
    $AutopilotDevices | ? serialNumber -eq $SignpostDevice.Serienummer | % {
        Write-Host "    Gevonden device: $($_.id) in levering $($SignpostDevice.Leveringsnummer)"
        Set-AutopilotDevice -displayName $SignpostDevice.Label -id $_.id -groupTag $SignpostDevice.Leveringsnummer
    }
}