$ClassGroups | % {
    $AzureADGroupName = $_.DisplayName
    $ClassName = $AzureADGroupName -replace 'romerocollege.*_',''
    Write-Host "Aanmaken van klas: $ClassName"

    $TeamName = "$YearCode$ClassName"
    $ClassGroup = Get-AzureADMSGroup -SearchString $TeamName | ? DisplayName -eq $TeamName
    Set-AzureAdMsGroup -ID $ClassGroup.Id -GroupTypes @("DynamicMembership", "Unified") -MembershipRuleProcessingState "On" -MembershipRule "(user.department -eq `"$($AzureADGroupName)`")"
}

