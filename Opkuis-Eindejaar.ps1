Connect-AzureAD -ErrorAction Stop
Connect-MicrosoftTeams

$CurrentYear = (Get-Date).Year % 100
$YearCode = "$($CurrentYear - 1)$($CurrentYear)_"
$OldYearCode = "$($CurrentYear - 4)$($CurrentYear - 3)_"

Write-Host "Archiveren van jaar $YearCode..."
Write-Host "Verwijderen van jaar $OldYearCode..."

# Klassen zoeken
$Classes = for($i = 1;$i -lt 8; $i++) { 
    Get-AzureADMsGroup -Filter "startswith(displayname,'$YearCode$i')" -All $true
}

# Klassen leegmaken
$Classes | % {
    $Class = $_
    Write-Host "Leegmaken klas $($Class.DisplayName)"

    $Team = Get-Team -DisplayName $Class.DisplayName 
    $GroupId = $Team.GroupId
<#    Get-TeamChannel -GroupId $GroupId | ? MembershipType -EQ Private | % {
        Remove-TeamChannel -GroupId $GroupId -DisplayName $_.DIsplayName
    }
    Set-AzureAdMsGroup -ID $_.Id -MembershipRuleProcessingState "Paused"
    #>
    Write-Host "    Team archiveren"
    Set-TeamArchivedState -GroupId $GroupId -Archived $true -SetSpoSiteReadOnlyForMembers $true
}


# Klassenraden leegmaken
Get-AzureADGroup -Filter "startswith(displayname,'$($YearCode)KR')" -All $true | % {
    $Klassenraad = $_
    Write-Host "Leegmaken klassenraad $($Klassenraad.DisplayName)"

    $Team = Get-Team -DisplayName $Klassenraad.DisplayName
    $GroupId = $Team.GroupId
    Get-TeamChannel -GroupId $GroupId | ? MembershipType -EQ Private | % {
        Remove-TeamChannel -GroupId $GroupId -DisplayName $_.DIsplayName
    }
    
    
    Write-Host "    Team archiveren"
    Set-TeamArchivedState -GroupId $GroupId  -Archived $true -SetSpoSiteReadOnlyForMembers $true
}

# Oude teams schrappen
Get-AzureADGroup -Filter "startswith(displayname,'$($OldYearCode)')" | % {
    Write-Host "Schrappen $($_.DisplayName)"
    Get-Team -DisplayName $_.DisplayName | Remove-Team
}
