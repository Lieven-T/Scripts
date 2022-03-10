$Users = Get-AzureADUser -Filter "Jobtitle eq 'leerling' and startswith(Department,'romerocollege_2')" -All $true
Write-Host "Oplijsten info voor $($Users.Count) gebruikers..."
$i = 0
$Users | % {
    Write-Progress -Activity "Toestellen zoeken..." -Status "$i/$($Users.Count) gedaan" -PercentComplete ($i / $Users.Count * 100)
    $i++
    $DeviceName = (Invoke-MSGraphRequest -Url "https://graph.microsoft.com/Beta/users/$($_.ObjectId)/managedDevices").value[0].deviceName

    $Props = @{
        Gebruiker = $_.UserPrincipalName
        Rol = $_.JobTitle
        Klas = ($_.Department -replace 'romerocollege.*_','')
        Toestelnaam = $DeviceName
    }
    New-Object -TypeName psobject -Property $Props
} | Export-Excel c:\temp\toestelinfo.xlsx