Get-ChildItem | ? ObjectClass -eq "organizationalUnit" | % {
    Write-Host $_.Name
    $ADLeerlingen = Get-ADUser -SearchBase $_.distinguishedName -Filter * -Properties emailaddress | select -ExpandProperty EmailAddress
    Get-AzureADGroup -Filter "displayName eq 'romerocollege_$($_.Name)'" | Get-AzureADGroupMember | ? { $_.UserPrincipalName -notin $ADLeerlingen } | % {
        if (($_.UserPrincipalName -split '@')[0].Length -gt 20) {
            Write-Host "    Lange naam: $($_.UserPrincipalName)"
        } else {
            Write-Host "    Ongeldige leerling: $($_.UserPrincipalName)"
            Remove-AzureADUser -ObjectId $_.ObjectId
        }
    }
}