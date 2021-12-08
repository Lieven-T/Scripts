if (-not (Get-InstalledModule windowsautopilotintune)) {
    Install-PackageProvider -Name "NuGet" -Force
    Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
    Install-Module windowsautopilotintune
}

$Credential = Get-Credential
Connect-AzureAD -Credential $Credential
Connect-MSGraph -Credential $Credential

$serialnumbers = Get-AutopilotDevice
$serials = Import-Excel C:\temp\devicelist.xlsx

foreach ($serial in $serials) {
    $sn = $serial.serial
    $devicename = $serial.devicename
    foreach ($serialnumber in $serialnumbers) {
        if ($serial.serial -eq $serialnumber.serialNumber)
        {$deviceids = Get-AutopilotDevice -serial $sn
        foreach ($deviceid in $deviceids)
    {

    Write-Host $deviceid.id

    Remove-AutopilotDevice -id $deviceid.id -displayName $devicename

    }

        }

        #else {Write-host "not ok"}

   

        }

    }

 

 

 