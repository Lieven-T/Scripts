﻿######################
### INITIALIZATION ###
######################

$FileLocation = "C:\Users\adweaver\Documents"
$YearCode = "2122_"

$Date = Get-Date
$TranscriptLocation = "$FileLocation\Sync-Teams_" + $Date.ToString("yyyyMMdd_HHmm") + ".log"
Start-Transcript $TranscriptLocation


$ClassGroups = Get-MgGroup -All -Filter "startswith(displayName,'romerocollege_')" | ? DisplayName -Match "romerocollege.*_[1-7]"
$ClassTeamGroups = Get-MgGroup -All -Filter "startswith(displayName,'$YearCode')" | ? DisplayName -Match "$YearCode[1-7]"
$ClassCodeList = $ClassTeamGroups | select -ExpandProperty DisplayName | % { $_ -replace $YearCode }

##################
### SYNC TEAMS ###
##################

$TeamsToCreate = @()
$ClassGroups | ? { ($_.DisplayName -replace "romerocollege.*_") -notin $ClassCodeList } | % {
    $ClassName = $AzureADGroupName -replace 'romerocollege.*_',''
    Write-Host "Aanmaken van klas: $ClassName"

    $TeamName = "$YearCode$ClassName"
    Write-Host "Aanmaken van team $_"
    $FunSettings = [PSCustomObject][Ordered]@{
        allowGiphy=$false
        allowStickersAndMemes= $false
        allowCustomMemes=$false
    }
    $Headers = [PSCustomObject][Ordered]@{"Content-Type"="application/json"}
    $Body = [PSCustomObject][Ordered]@{
        "template@odata.bind"="https://graph.microsoft.com/beta/teamsTemplates('educationClass')"
        displayName=$TeamName
        description=$TeamName
        mailNickName=$TeamName
        funSettings=$FunSettings
    }
    $TeamsToCreate += [PSCustomObject][Ordered]@{
        Id=$_
        Method='POST'
        Url="/teams"
        Headers=$Headers
        Body=$Body
    }
}

$InitialResponses = @()
for($i=0;$i -lt $TeamsToCreate.count;$i+=20) {
    Write-Progress -Activity "Teams aanmaken..." -Status "$i/$($TeamsToCreate.Count) gedaan" -PercentComplete ($i / $TeamsToCreate.Count * 100)
    $Request = @{}           
    $Request['requests'] = ($TeamsToCreate[$i..($i+19)])
    $RequestBody = $Request | ConvertTo-Json -Depth 4
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "202" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
    $InitialResponses += $Response.responses
}

$Actions = @()
$InitialResponses | % {
    $Actions += @([PSCustomObject][Ordered]@{
        Id=$_.id
        Method='GET'
        Url=$_.headers.Location
    })
}

while($Actions.count) {
    Write-Host "Wachten op $($Actions.count)..."
    Start-Sleep 45
    $Responses = @()
    for($i=0;$i -lt $Actions.count;$i+=20) {
        Write-Progress -Activity "Opvragen van $($Actions.Count) acties" -Status "$i/$($Actions.Count) gedaan" -PercentComplete ($i / $Actions.Count * 100)
        $Request = @{}           
        $Request['requests'] = ($Actions[$i..($i+19)])
        $RequestBody = $Request | ConvertTo-Json -Depth 4
        $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
        $Response.responses | ? status -ne "200" | % {
            Write-Error "Probleem met $($_.id): $($_.body.error.message)"
        }
        $Responses += $Response.responses | ? { $_.status -eq "200" -and $Response.responses[0].body.status -ne "succeeded" }
    }
    $Actions = $Actions | ? Id -in $Responses.Id
}

$StaticClassGroups = Get-MgGroup -Filter "startswith(displayname,'$($Yearcode)')" -all | ? { $_.DisplayName -match "romerocollege.*_[1-7]" -and 'DynamicMembership' -notin $_.GroupTypes }
$StaticClassGroups | % {
    $OriginGroup = $ClassGroups | ? { $_.DisplayName -match "romerocollege.*_$(($_.DisplayName -split '_')[-1])" }
    Write-Host "Instellen van $($_.DisplayName) naar lidmaatschap van $($OriginGroup.DisplayName)"
    Update-MgGroup -GroupId "5242dd38-36dd-4fb7-bfaf-d8b3f0d7c5fc" -BodyParameter @{
        GroupTypes=@("DynamicMembership","Unified")
        MembershipRuleProcessingState="On"
        MembershipRule="(user.department -eq `"$($OriginGroup.DisplayName)`")"
    }
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