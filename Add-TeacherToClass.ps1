Param
(
    [Parameter(Mandatory=$True,Position=1)]
    [string]$Teacher,
    [Parameter(Mandatory=$True,Position=2)]
    [Array]$ClassList

)

try {
    Get-AzureADTenantDetail | Out-Null
} catch {
    $Credentials = Get-Credential
    Connect-AzureAD -ErrorAction Stop -Credential $Credentials
    Connect-MicrosoftTeams -Credential $Credentials
}

[string]$YearCode = "2122_"

# Leraar toevoegen aan klas
$User = Get-AzureADUser -Filter "UserPrincipalName eq '$Teacher@romerocollege.be'"
$ClassList | % {
    Write-Host "Toevoegen van $Teacher aan klas $YearCode$_"
    $Team = Get-Team -DisplayName "$YearCode$_"
    Add-TeamUser -GroupId $Team.GroupId -User $User.ObjectId -Role Owner

    Write-Host "Toevoegen van $Teacher aan klassenraad $($YearCode)KR_$_"
    $Group = Get-AzureADGroup -SearchString "$($YearCode)KR_$_"
    Add-AzureADGroupOwner -ObjectId $Group.ObjectId -RefObjectId $User.ObjectId
}