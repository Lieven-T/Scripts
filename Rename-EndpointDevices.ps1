Param
(
    [Parameter(Position=1,ValueFromPipeline=$true)]
    [ValidateScript({Test-Path $_})] 
    [string]$InputFile
)

if (-not (Get-InstalledModule Microsoft.Graph.Intune)) {
    Install-PackageProvider -Name "NuGet" -Force
    Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
    Install-Module Microsoft.Graph.Intune
}

if (-not (Get-InstalledModule windowsautopilotintune)) {
    Install-PackageProvider -Name "NuGet" -Force
    Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
    Install-Module windowsautopilotintune
}

$FileList = if ($input) {
    $input | select -ExpandProperty FullName
} elseif ($InputFile) {
    $InputFile
} else {
    throw "No data found"
}

Write-Host "`n==========`nFETCH DATA`n==========`n"
Connect-MSGraph
$Devices = Get-IntuneManagedDevice | Get-MSGraphAllPages
$AutopilotDevices = Get-AutopilotDevice

# TODO: batching
foreach ($i in $FileList) {
    $InputFile = $i
    $InputFile -match "SP.*S\d" | Out-Null
    $Leveringsnummer = $Matches[0]
    
    $Msg = "LEVERINGSNUMMER: $Leveringsnummer"
    $Delimiter = "=" * $Msg.Length
    Write-Host $Delimiter
    Write-Host $Msg
    Write-Host $Delimiter

    Write-Host "`nENDPOINT`n--------"
    Import-Excel -Path $InputFile | % {
        $SignpostSerialNumber = $_.'Serienummer van apparaat'
        $SignpostHostname = $_.Volgnummer
        $Device = $Devices | ? serialNumber -eq $SignpostSerialNumber
        if ($Device.deviceName -eq $SignpostHostname) {
            return
        }

        Write-Host "Toestel $SignpostHostname met serienummer $SignpostSerialNumber"
        if (-not $Device) {
            Write-Host "    Toestel niet gevonden"
            return
        }
        if ($Device.count -gt 1) {
            Write-Host "    Meerdere toestellen gevonden"
            return
        }
        Write-Host "    $($Device.deviceName) -> $SignpostHostname"
        Invoke-MSGraphRequest -HttpMethod POST -Url "https://graph.microsoft.com/Beta/deviceManagement/managedDevices/$($Device.id)/setDeviceName" -Content "{ deviceName:`"$SignpostHostname`" }"
    }

    Write-Host "`nAUTOPILOT`n---------"
    Import-Excel -Path $InputFile | % {
        $SignpostSerialNumber = $_.'Serienummer van apparaat'
        $SignpostHostname = $_.Volgnummer
    
        $Device = $AutopilotDevices | ? serialNumber -eq $SignpostSerialNumber 
        if (($Device.displayName -eq $SignpostHostname) -and ($Device.groupTag -eq $Leveringsnummer)) {
            return
        }

        Write-Host "Toestel $SignpostHostname met serienummer $SignpostSerialNumber"
        if (-not $Device) {
            Write-Host "    Toestel niet gevonden"
            return
        }
        if ($Device.count -gt 1) {
            Write-Host "    Meerdere toestellen gevonden"
            return
        }

        Write-Host "    $($Device.displayName) -> $SignpostHostname"
        Write-Host "    $($Device.groupTag) -> $Leveringsnummer"
        Invoke-MSGraphRequest -HttpMethod POST -Url "https://graph.microsoft.com/Beta/deviceManagement/windowsAutopilotDeviceIdentities/$($Device.id)/UpdateDeviceProperties" -Content "{ displayName:`"$SignpostHostname`", groupTag: `"$Leveringsnummer`" }"
    }
}
