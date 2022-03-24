Param
(
    [Parameter(Position=1,ValueFromPipeline=$true)]
    [ValidateScript({Test-Path $_})] 
    [string]$InputFile
)

Connect-Graph -Scopes @("DeviceManagementManagedDevices.ReadWrite.All","DeviceManagementManagedDevices.PrivilegedOperations.All")
$ExcelData = Import-Excel $InputFile | ? { $_.'Toestel ID' -and $_.'Toestel ID' -is [String]} | Sort-Object -Property Klas

Write-Host "Opvragen status van $($ExcelData.Count) laptops..."
$DevicesToQuery = @()
$ExcelData | % {
    Write-Host "$($_.'Toestel ID'): $($_.toestel) van $($_.Gebruiker) uit $($_.Klas)"                                                                                                                                          
    $DevicesToQuery += [PSCustomObject][Ordered]@{
        Id=$_.'Toestel ID'
        Method='GET'
        Url="/deviceManagement/managedDevices/$($_.'Toestel ID')"
    }
}

$DeviceData = @()
for($i=0;$i -lt $DevicesToQuery.count;$i+=20){                                                                                                                                              
    Write-Progress -Activity "Toestelstatus opvragen..." -Status "$i/$($DevicesToQuery.Count) gedaan" -PercentComplete ($i / $DevicesToQuery.Count * 100)
    $Request = @{}                
    $Request['requests'] = ($DevicesToQuery[$i..($i+19)])
    $RequestBody = $Request | ConvertTo-Json -Depth 3
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "200" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
    $Response.responses | ? status -eq "200" | % {
        $Body = $_.body
        $DeviceData += [PSCustomObject]@{
            'Toestel ID' = $Body.id
            Toestel = $Body.deviceName
            Status = $Body.managementState
        }
    }
}

$DeviceData | Export-Excel C:\temp\Toestelstatus.xlsx
