Connect-Graph -Scopes @("User.ReadBasic.All","User.Read.All","DeviceManagementManagedDevices.Read.All","Directory.Read.All")
$Users = Get-MgUser -Filter "Jobtitle eq 'leerling' and startswith(Department,'romerocollege_2')" -All -Property @("Department","Id","UserPrincipalName","JobTitle")

Write-Host "Oplijsten info voor $($Users.Count) gebruikers..."
$Output = for($i=0;$i -lt $Users.count;$i+=20){                                                                                                                                              
    Write-Progress -Activity "Toestellen zoeken..." -Status "$i/$($Users.Count) gedaan" -PercentComplete ($i / $Users.Count * 100)
    $Request = @{}             
    $Request['requests'] = $Users[$i..($i+19)] | % {
        [PSCustomObject][Ordered]@{
            id=$_.UserPrincipalName
            method='GET'
            Url="/users/$($_.id)/managedDevices"
        }
    }
    $RequestBody = $Request | ConvertTo-Json -Depth 3
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "200" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
    $Users[$i..($i+19)] | % {
        $Device = ($Response.responses | ? id -eq $_.UserPrincipalName).body.value
        [PSCustomObject][Ordered]@{
            Gebruiker=$_.UserPrincipalName
            Klas=($_.Department -replace 'romerocollege.*_','')
            Rol=$_.JobTitle
            Toestel=$Device.deviceName
            'Toestel ID'=$Device.id
        }
    }
}
$Output | Export-Excel c:\temp\toestelinfo.xlsx