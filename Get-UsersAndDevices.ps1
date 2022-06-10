Connect-Graph -Scopes @("User.ReadBasic.All","User.Read.All","DeviceManagementManagedDevices.Read.All","Directory.Read.All")
$Users = Get-MgUser -Filter "startswith(Department,'romerocollege_')" -All -Property @("Department","Id","UserPrincipalName","JobTitle") | ? Department -Match 'romerocollege_3'

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
        if ($Device -is [Array]) {
                $Rol = $_.JobTitle
                $Gebruiker = $_.UserPrincipalName
                $Klas = ($_.Department -replace 'romerocollege.*_','')
                if (-not $Device.Count) {
                    [PSCustomObject][Ordered]@{
                        Gebruiker = $Gebruiker
                        Klas = $Klas
                        Rol = $Rol
                        Toestel = ''
                        'Toestel ID'= ''
                        Model = ''
                        EnrolledDateTime = ''
                        LastSyncDateTime = ''
                    }
                }
                $Device | % {
                    [PSCustomObject][Ordered]@{
                        Gebruiker = $Gebruiker
                        Klas = $Klas
                        Rol = $Rol
                        Toestel = $_.deviceName
                        'Toestel ID'= $_.Id
                        Model = $_.Model
                        EnrolledDateTime = $_.enrolledDateTime
                        LastSyncDateTime = $_.lastSyncDateTime
                    }
                }
        } else {
            [PSCustomObject][Ordered]@{
                Gebruiker=$_.UserPrincipalName
                Klas=($_.Department -replace 'romerocollege.*_','')
                Rol=$_.JobTitle
                Toestel=$Device.deviceName
                'Toestel ID'=$Device.id
            }
        }
    }
}
$Output | Sort -Property Klas,Gebruiker | Export-Excel c:\temp\toestelinfo.xlsx