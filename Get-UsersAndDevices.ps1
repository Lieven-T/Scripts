Connect-Graph -Scopes @("User.ReadBasic.All","User.Read.All","DeviceManagementManagedDevices.Read.All","Directory.Read.All")
$Users = Get-MgUser -Filter "Jobtitle eq 'leerling' and startswith(Department,'romerocollege_2')" -All -Property @("Department","Id","UserPrincipalName","JobTitle")
Write-Host "Oplijsten info voor $($Users.Count) gebruikers..."

$Output = for($i=0;$i -lt $Users.count;$i+=20){                                                                                                                                              
    Write-Progress -Activity "Toestellen zoeken..." -Status "$i/$($Users.Count) gedaan" -PercentComplete ($i / $Users.Count * 100)
    $Request = @{}                
    $Request['requests'] = ($Users[$i..($i+19)] | select @{n='id';e={$_.id}},@{n='method';e={'GET'}},`
        @{n='Url';e={"/users/$($_.id)/managedDevices"}})
    $RequestBody = $Request | ConvertTo-Json -Depth 3
    $Responses = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    ($Users[$i..($i+19)] | select @{n='Gebruiker';e={$_.UserPrincipalName}},@{n='Klas';e={($_.Department -replace 'romerocollege.*_','')}},`
        @{n='Rol';e={$_.JobTitle}},@{n='Toestel';e={($Responses.responses | ? id -eq $_.Id).body.value.deviceName}},@{n='Toestel ID';e={($Responses.responses | ? id -eq $_.Id).body.value.id}})
}
$Output | Export-Excel c:\temp\toestelinfo.xlsx