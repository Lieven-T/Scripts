Param
(
    [Array]$ClassList
)

try {
    Get-AzureADTenantDetail | Out-Null
} catch {
    $Credentials = Get-Credential
    Connect-AzureAD -ErrorAction Stop -Credential $Credentials
    Connect-MicrosoftTeams -Credential $Credentials
}

[string]$YearCode = "2122"
$TeamName = "$($YearCode)"
$ClassList | % {
    $TeamName += "_$_"
}

Write-Host "Aanmaken team $TeamName"
$NewTeam = New-Team -DisplayName $TeamName -MailNickName $TeamName -Template EDU_Class

$ClassList | % {
    $Group = Get-AzureADGroup -Filter "displayname eq '$($YearCode)_$($_)'"
    Write-Host "Toevoegen leerlingen uit $_"
    Get-AzureADGroupMember -ObjectId $Group.ObjectId | % {
        Write-Host "    $($_.UserPrincipalName)"
        Add-TeamUser -GroupId $NewTeam.GroupId -User $_.UserPrincipalName -Role Member
    }
    Write-Host "Toevoegen leerkrachten uit $_"
    Get-AzureADGroupOwner -ObjectId $Group.ObjectId | % {
        Write-Host "    $($_.UserPrincipalName)"
        Add-TeamUser -GroupId $NewTeam.GroupId -User $_.UserPrincipalName -Role Owner
    }
}
Write-Host "Schrappen mezelf"
Remove-TeamUser -GroupId $NewTeam.GroupId -User "lieven.tronckoe@romerocollege.be"