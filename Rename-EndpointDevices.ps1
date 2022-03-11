Param
(
    [Parameter(Position=1,ValueFromPipeline=$true)]
    [ValidateScript({Test-Path $_})] 
    [string]$InputFile
)

if (-not (Get-InstalledModule Microsoft.Graph)) {
    Install-PackageProvider -Name "NuGet" -Force
    Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
    Install-Module Microsoft.Graph
}

$FileList = if ($input) {
    $input | select -ExpandProperty FullName
} elseif ($InputFile) {
    $InputFile
} else {
    throw "No data found"
}

Write-Host "`n==========`nFETCH DATA`n==========`n"
Connect-Graph -Scopes @("DeviceManagementManagedDevices.ReadWrite.All","DeviceManagementServiceConfig.ReadWrite.All","DeviceManagementManagedDevices.PrivilegedOperations.All")
Select-MgProfile -Name "beta"
$Response = Invoke-GraphRequest -Uri "https://graph.microsoft.com/Beta/deviceManagement/managedDevices"
$EndpointDevices = $Response.value
while ($Response.'@odata.nextLink') {
    $Response = Invoke-GraphRequest -Uri $Response.'@odata.nextLink'
    $EndpointDevices += $Response.value
}
$EndpointSerialNumbers = $EndpointDevices | select -ExpandProperty "serialNumber"


$Response = Invoke-GraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeviceIdentities"
$AutopilotDevices = $Response.value
while ($Response.'@odata.nextLink') {
    $Response = Invoke-GraphRequest -Uri $Response.'@odata.nextLink'
    $AutopilotDevices += $Response.value
}
$AutopilotSerialNumbers = $EndpointDevices | select -ExpandProperty "serialNumber"

foreach ($i in $FileList) {
    $InputFile = $i
    $InputFile -match "SP.*S\d" | Out-Null
    $Leveringsnummer = $Matches[0]
    
    $Msg = "LEVERINGSNUMMER: $Leveringsnummer"
    $Delimiter = "=" * $Msg.Length
    Write-Host $Delimiter
    Write-Host $Msg
    Write-Host $Delimiter

    $ExcelData = Import-Excel -Path $InputFile
    Write-Host "`nENDPOINT`n--------"
    $DevicesToRename = @()
    $i=0
    $ExcelData | % {
        $EndpointDevice = $EndpointDevices | ? serialNumber -eq $_.'Serienummer van apparaat'
        if (-not $EndpointDevice) { 
            Write-Host "Niet gevonden: toestel $($_.Volgnummer) met serienummer $($_.'Serienummer van apparaat')"
            return
        }
        if ($EndpointDevice -is [Array]) {
            Write-Host "Meerdere toestellen gevonden voor toestel $($_.Volgnummer) met serienummer $($_.'Serienummer van apparaat')"
            return
        }

        $SignpostDeviceName = $_.Volgnummer
        if ($EndpointDevice.deviceName -ne $SignpostDeviceName) {
            Write-Host "$($EndpointDevice.deviceName) -> $SignpostDeviceName"
            $Headers = [PSCustomObject][Ordered]@{"Content-Type"="application/json"}
            $Body = [PSCustomObject][Ordered]@{ deviceName=$SignpostDeviceName}
            $DevicesToRename += [PSCustomObject][Ordered]@{
                id=$i++
                Method="POST"
                Url="deviceManagement/managedDevices/$($EndpointDevice.id)/setDeviceName"
                Headers=$Headers
                Body=$Body
             }
        }
    }

    Write-Host "Request verzenden..."
    for($i=0;$i -lt $DevicesToRename.count;$i+=20){                                                                                                                                              
        $Request = @{}                
        $Request['requests'] = ($DevicesToRename[$i..($i+19)])
        $RequestBody = $Request | ConvertTo-Json -Depth 3
        $Responses = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    }

    Write-Host "`nAUTOPILOT`n---------"
    $DevicesToRename = @()
    $i=0
    $ExcelData | % {
        $AutopilotDevice = $AutopilotDevices | ? serialNumber -eq $_.'Serienummer van apparaat'
        if (-not $AutopilotDevice) { 
            Write-Host "Niet gevonden: toestel $($_.Volgnummer) met serienummer $($_.'Serienummer van apparaat')"
            return
        }
        if ($EndpointDevice -is [Array]) {
            Write-Host "Meerdere toestellen gevonden voor toestel $($_.Volgnummer) met serienummer $($_.'Serienummer van apparaat')"
            return
        }

        $SignpostDeviceName = $_.Volgnummer
        if ($AutopilotDevice.displayName -ne $SignpostDeviceName -or $AutopilotDevice.groupTag -ne $Leveringsnummer) {
            Write-Host "Toestel $($SignpostDeviceName):`n    $($AutopilotDevice.displayName) -> $SignpostDeviceName`n    $($AutopilotDevice.groupTag) -> $Leveringsnummer"
            $Headers = [PSCustomObject][Ordered]@{"Content-Type"="application/json"}
            $Body = [PSCustomObject][Ordered]@{ 
                displayName=$SignpostDeviceName
                groupTag=$Leveringsnummer
            }
            $DevicesToRename += [PSCustomObject][Ordered]@{
                id=$i++
                Method="POST"
                Url="deviceManagement/windowsAutopilotDeviceIdentities/$($AutopilotDevice.id)/updateDeviceProperties"
                Headers=$Headers
                Body=$Body
             }
        }
    }

    Write-Host "Request verzenden..."
    for($i=0;$i -lt $DevicesToRename.count;$i+=20){                                                                                                                                              
        $Request = @{}                
        $Request['requests'] = ($DevicesToRename[$i..($i+19)])
        $RequestBody = $Request | ConvertTo-Json -Depth 3
        $Responses = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    }

}
