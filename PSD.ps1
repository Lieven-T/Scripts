$File = "PSD_Teams.xlsx"
$UserColumn = "Gebruiker"
$GroupColumn = "Team"
$RoleColumn = "Rol"

Connect-AzureAD

$AllGroups = Get-AzureADGroup -SearchString "2021_PSD"

$DataRaw = Import-Excel -Path $File
$Data = $DataRaw | ? { $_.$UserColumn -match "@romerocollege" }

$ExcelGroupNames = $Data | select -ExpandProperty $GroupColumn -Unique 
$ADGroupNames = $AllGroups | select -ExpandProperty DisplayName -Unique 

$GroupsNotInAD = $GroupNames | ? { $_ -notin $ADGroupNames }
$GroupsNotInExcel = $ADGroupNames | ? { $_ -notin $ExcelGroupNames }
Write-Host "Groepen niet in AD:"
Write-Host $GroupsNotInAD
Write-Host "Groepen niet in Excel:"
Write-Host $GroupsNotInExcel

# Leden toevoegen
$Data | % {
    $Group = $AllGroups | ? DisplayName -EQ $_.$GroupColumn
    $User = Get-AzureADUser -SearchString ($_.Gebruiker -split "@")[0]
    if ($User) {
        Write-Host "Toevoegen $($User.UserPrincipalName) aan $($Group.DisplayName) als $($_.$RoleColumn)"
        if ($_.$RoleColumn -eq "eigenaar") {
            Add-AzureADGroupOwner -ObjectId $Group.ObjectId -RefObjectId $User.ObjectId
        } else {
            Add-AzureADGroupMember -ObjectId $Group.ObjectId -RefObjectId $User.ObjectId
        }
    } else {
        Write-Host "Gebruiker niet gevonden: $($_.$UserColumn)" -ForegroundColor Red
    }
}

# Leden schrappen
$AllGroups | % {
    $Group = $_
    Get-AzureADGroupMember -ObjectId $_.ObjectId | ? { $_.UserPrincipalName -notin ($Data | ? $GroupColumn -NE $Group.DisplayName) } | % {
        Write-Host "Schrappen $($_.UserPrincipalName) uit $($Group.DisplayName)"
        # Remove-AzureADGroupMember -ObjectId $Group.ObjectId -MemberId $_.ObjectId
    }
}