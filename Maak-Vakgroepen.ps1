try {
    Get-AzureADTenantDetail | Out-Null
} catch {
    $Credentials = Get-Credential
    Connect-AzureAD -ErrorAction Stop -Credential $Credentials
    Connect-MicrosoftTeams -Credential $Credentials
}

$Vakgroepen = Import-Excel .\vakgroepen1GR.xlsx

$vakgroepen | select -ExpandProperty vakgroep -Unique | % {
    Write-Host "Aanmaken vakgroep $_"
    New-Team -DisplayName $_ -MailNickName $_ -AllowCreatePrivateChannels $false
}

$vakgroepen | % {
    Write-Host "Toevoegen van $($_.leraar) aan $($_.vakgroep) als $($_.rol)"
    if (Get-AzureADUser -SearchString $_.leraar) {
        $UserName = "$($_.leraar)@romerocollege.be"
        if ($_.rol -eq "eigenaar") {
            Get-Team -DisplayName $_.vakgroep | Add-TeamUser -User $UserName -Role Owner
        } else {
            Get-Team -DisplayName $_.vakgroep | Add-TeamUser -User $UserName -Role Member
        }
    } else {
        Write-Host "    Gebruiker niet gevonden"
    }
}

$vakgroepen | select -ExpandProperty vakgroep -Unique | % {
    Write-Host "Schrappen Lieven uit $_"
    Get-Team -DisplayName $_ | Remove-TeamUser -User "lieven.tronckoe@romerocollege.be"
}
