Connect-Graph -Scopes @("Directory.AccessAsUser.All")
$AllDevices = Get-MgDevice -All
$CutoffDate = (Get-Date).AddDays(-200)

$DevicesToDelete = @()
$AllDevices | ? { $_.ApproximateLastSignInDateTime -lt $CutoffDate -and $_.TrustType -ne "AzureAd" } | % {
    Write-Host "Device wissen: $($_.DisplayName)"
    $DevicesToDelete += [PSCustomObject][Ordered]@{
        Id=$_.Id
        Method='DELETE'
        Url="/devices/$($_.Id)"
    }
}

$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

$choiceRTN = $host.UI.PromptForChoice("WISSEN TOESTELLEN", "Wissen van $($DevicesToDelete.Count)", $options, 1)
if ( $choiceRTN -eq 1 ) {
    return
}

for($i=0;$i -lt $DevicesToDelete.count;$i+=20){                                                                                                                                              
    Write-Progress -Activity "Toestellen wissen..." -Status "$i/$($DevicesToDelete.Count) gedaan" -PercentComplete ($i / $DevicesToDelete.Count * 100)
    $Request = @{}                
    $Request['requests'] = ($DevicesToDelete[$i..($i+19)])
    $RequestBody = $Request | ConvertTo-Json -Depth 3
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "204" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }

}
