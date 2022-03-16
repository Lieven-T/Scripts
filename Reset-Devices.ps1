Param
(
    [Parameter(Position=1,ValueFromPipeline=$true)]
    [ValidateScript({Test-Path $_})] 
    [string]$InputFile
)

Connect-Graph -Scopes @("DeviceManagementManagedDevices.ReadWrite.All","DeviceManagementManagedDevices.PrivilegedOperations.All")
$ExcelData = Import-Excel $InputFile | ? { $_.'Toestel ID' -and $_.'Toestel ID' -is [String]} | Sort-Object -Property Klas

Write-Host "Resetten van $($ExcelData.Count) laptops..."
$DevicesToReset = @()
$ExcelData | % {
    Write-Host "$($_.'Toestel ID'): $($_.toestel) van $($_.Gebruiker) uit $($_.Klas)"                                                                                                                                          
    $DevicesToReset += [PSCustomObject][Ordered]@{
        Id=$_.Toestel
        Method='POST'
        Url="/managedDevices/$($_.'Toestel ID')/wipe"
    }
}

$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

$choiceRTN = $host.UI.PromptForChoice("RESETTEN TOESTELLEN", "Resetten van $($DevicesToReset.Count)", $options, 1)
if ( $choiceRTN -eq 1 ) {
    return
}

for($i=0;$i -lt $DevicesToReset.count;$i+=20){                                                                                                                                              
    Write-Progress -Activity "Toestellen resetten..." -Status "$i/$($DevicesToReset.Count) gedaan" -PercentComplete ($i / $DevicesToReset.Count * 100)
    $Request = @{}                
    $Request['requests'] = ($DevicesToReset[$i..($i+19)])
    $RequestBody = $Request | ConvertTo-Json -Depth 3
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "204" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }

}
