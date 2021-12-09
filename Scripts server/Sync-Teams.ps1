######################
### INITIALIZATION ###
######################

$FileLocation = "C:\Users\adweaver\Documents"
$YearCode = "2122_"

$Date = Get-Date
$TranscriptLocation = "$FileLocation\Sync-Teams_" + $Date.ToString("yyyyMMdd_HHmm") + ".log"
Start-Transcript $TranscriptLocation

$Password = Get-Content "$FileLocation\cred.txt" | convertto-securestring
$Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "adweaver@romerocollege.be",$Password
Import-Module AzureADPreview
Connect-AzureAD -ErrorAction Stop -Credential $Credentials
Connect-MicrosoftTeams -Credential $Credentials
Connect-ExchangeOnline -Credential $Credentials -ShowBanner:$false

$ClassGroups = for($i = 1;$i -lt 8; $i++) { 
    Get-AzureADMsGroup -Filter "startswith(displayname,'romerocollege_$i') or startswith(displayname,'romerocollege_cvw_$i')" -All $true
}

$ClassTeamGroups = for($i = 1;$i -lt 8; $i++) { 
    Get-AzureADMsGroup -Filter "startswith(displayname,'$YearCode$i')" -All $true
}

$ClassCodeList = $ClassTeamGroups | select -ExpandProperty DisplayName | % { $_ -replace $YearCode }

##################
### SYNC TEAMS ###
##################

$ClassGroups | ? { ($_.DisplayName -replace "romerocollege.*_") -notin $ClassCodeList } | % {
    $AzureADGroupName = $_.DisplayName
    $ClassName = $AzureADGroupName -replace 'romerocollege.*_',''
    Write-Host "Aanmaken van klas: $ClassName"

    $TeamName = "$YearCode$ClassName"
    $ClassTeam = Get-Team -DisplayName $TeamName | ? DisplayName -eq $TeamName
    if (-not $ClassTeam -and (Get-Team -MailNickName $TeamName)) { 
        Write-Host "    Naam van klas is veranderd, wordt overgeslagen"
    } else {
        if (-not $ClassTeam) {
            Write-Host "    Klasteam bestaat nog niet: aanmaken"
            $ClassTeam = New-Team -DisplayName $TeamName -MailNickName "$YearCode$($ClassName -replace " ","_")" -Template EDU_Class -AllowGiphy $false -AllowStickersAndMemes $false -AllowCustomMemes $false
            $ClassGroup = Get-AzureADMSGroup -SearchString $TeamName | ? DisplayName -eq $TeamName
            Set-AzureAdMsGroup -ID $ClassGroup.Id -GroupTypes @("DynamicMembership", "Unified") -MembershipRuleProcessingState "On" -MembershipRule "(user.department -eq `"$($AzureADGroupName)`")"
        } else {
            Write-Host "    Team gevonden"
            $ClassGroup = Get-AzureADMSGroup -SearchString $TeamName | ? DisplayName -eq $TeamName
            Set-AzureAdMsGroup -ID $ClassGroup.Id -GroupTypes @("DynamicMembership", "Unified") -MembershipRuleProcessingState "On" -MembershipRule "(user.department -eq `"$($AzureADGroupName)`")"
        }

        # Distributielijst syncen
        $DistGroupName = "leerlingen.$($ClassName.ToLower())"
        $DistGroupDispName = "Leerlingen $ClassName"
        $DistGroup = Get-DynamicDistributionGroup -Identity $DistGroupName -ErrorAction SilentlyContinue
        if (-not $DistGroup) {
            Write-Host "    Distributielijst bestaat nog niet: aanmaken"
            $DistGroup = New-DynamicDistributionGroup -DisplayName $DistGroupDispName -Name $DistGroupName -ConditionalDepartment $AzureADGroupName -IncludedRecipients "MailboxUsers" 
        } else {
            Write-Host "    Distributielijst gevonden"
        }
    }
}

# Filter obsolete mailing lists
$ClassNameList = $ClassGroups | select -ExpandProperty DisplayName | % { $_.ToLower() -replace 'romerocollege.*_','' } 
Get-DynamicDistributionGroup | ? {$_.Name -match "^leerlingen\." -and ($_.Name -replace "^leerlingen\.") -notin $ClassNameList } | % {
    Write-Host "Schrappen distributielijst $($_.Name)"
    Remove-DynamicDistributionGroup -Identity $_.Id -Confirm:$false
}


###############
### SYNC SP ###
###############

$SharedDocsName = "Gedeelde documenten"
$StudentDocs = "Documenten Leerlingen"

Import-Excel "$FileLocation\teams.xlsx" | Select -ExpandProperty Klas -Unique | % {
    $Class = $_
    $ClassTeam = "$YearCode$_"
    Write-Host "Verwerken klas $Class"
    $Team = $null
    $Team = Get-Team -DisplayName $ClassTeam -ErrorAction Stop
    $SiteUrl = (Get-UnifiedGroup -Identity $Team.GroupId).SharepointSiteUrl
    $SharedDocs = (Get-UnifiedGroup -Identity $Team.GroupId).SharePointDocumentsUrl
    try {
        Connect-PnPOnline -Url $SiteUrl -Credentials $Credentials -ErrorAction Stop
    } catch {
        Write-Host "    Toegang geweigerd, eigenaar toevoegen"
        Set-PnPTenantSite -Url $SiteUrl  -Owners "adweaver@romerocollege.be"
        Connect-PnPOnline -Url $SiteUrl -Credentials $Credentials -ErrorAction Stop
    }
    $Roles = Get-PnPRoleDefinition
    $EditRole = ($roles | ? RoleTypeKind -eq "Contributor").Name
    $ReadRole = ($roles | ? RoleTypeKind -eq "Reader").Name
    $Members  = Get-PnPGroup -AssociatedMemberGroup
    $Users = Get-AzureADGroup -SearchString $ClassTeam | Get-AzureADGroupMember | ? UserPrincipalName -Match "@student"

    $Channel = Get-TeamChannel -GroupId $Team.GroupId | ? DisplayName -EQ $StudentDocs
    if (-not $Channel) {
        Write-Host "    Kanaal aanmaken..."
        $Channel = New-TeamChannel -GroupId $Team.GroupId -DisplayName $StudentDocs -ErrorAction Stop
        Set-PnPFolderPermission -List $SharedDocs -Identity "$SharedDocs/$StudentDocs" -Group $Members -AddRole $ReadRole -ClearExisting
    }
    $SubFolders = Get-PnPFolderItem -FolderSiteRelativeUrl "$SharedDocsName/$StudentDocs"
    $SubFolderNames = $Subfolders | select -ExpandProperty Name
    $UserNames = $Users | select -ExpandProperty UserPrincipalName | % { ($_ -split "@")[0] }
    $Users | ? { ($_.UserPrincipalName -split "@")[0] -notin ($SubFolderNames) } | % {
        $UserName = ($_.UserPrincipalName -split "@")[0]
        Write-Host "    Aanmaken map $UserName"
        Add-PnPFolder -Name $UserName -Folder $SharedDocsName/$StudentDocs
        Set-PnPFolderPermission -List $SharedDocsName/$StudentDocs -Identity "$SharedDocsName/$StudentDocs/$UserName" -User $_.UserPrincipalName -AddRole $EditRole -ClearExisting
    }
    $Subfolders | ? Name -NotIn $UserNames | % {
        Write-Host "    Verwijderen $($_.Name)"
        Remove-PnPFolder -Name $_.Name -Folder $SharedDocsName/$StudentDocs -Recycle -Force
    }
}

Stop-Transcript
Send-MailMessage -From 'Server Alerter CVD <alerter-cvd@romerocollege.be>' -To 'it-cvd@romerocollege.be' -Subject 'Sync Teams' -Attachments $TranscriptLocation -SmtpServer "romerocollege-be.mail.protection.outlook.com"