Get-ChildItem "\\campus\ll" | % {
    Write-Host "Opkuisen klas $_"
    Get-ChildItem  $_.FullName | % {
        Write-Host "   $_"
        Get-ChildItem $_.FullName | Remove-Item -Recurse -Force
    }
}

$PassWord = ConvertTo-SecureString -AsPlainText -String "RomeroGeheim!" -Force
Get-ADUser -SearchBase "OU=leerlingen,OU=gebruikers,OU=school,DC=cnn,DC=romerocollege,DC=edu" -Filter * | % {
    Write-Host "Reser gebruiker $($_.Name)"
    Set-ADAccountPassword $_ -NewPassword $Password -Reset
    Set-ADUser $_ -ChangePasswordAtLogon $true
}