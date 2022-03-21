Write-Error "DEBUGGEN!"
return

$CurrentYear = (Get-Date).Year % 100
$YearCode = "$($CurrentYear - 1)$($CurrentYear)_"
$OldYearCode = "$($CurrentYear - 4)$($CurrentYear - 3)_"

$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

$choiceRTN = $host.UI.PromptForChoice("OPKUIS EINDEJAAR", "Archiveren van jaar $YearCode en verwijderen van jaar $OldYearCode", $options, 1)
if ( $choiceRTN -eq 1 ) {
    return
}

# Klassen zoeken
Connect-Graph -Scopes @("Group.ReadWrite.All","TeamSettings.ReadWrite.All")
$Response = Invoke-GraphRequest -Uri "https://graph.microsoft.com/beta/teams?`$filter=startswith(displayName,'$($YearCode)')"
$Teams = $Response.value
while ($Response.'@odata.nextLink') {
    $Response = Invoke-GraphRequest -Uri $Response.'@odata.nextLink'
    $Teams += $Response.value
}

# LIST CHANNELS
$TeamChannels = @()
for($i=0;$i -lt $Teams.count;$i+=20){                                                                                                                                              
    Write-Progress -Activity "Privékanalen zoeken..." -Status "$i/$($Teams.Count) gedaan" -PercentComplete ($i / $Teams.Count * 100)
    $Request = @{}
    $Request.requests = $Teams[$i..($i+19)] | % {
        [PSCustomObject][Ordered]@{
            id=$_.ID
            method='GET'
            Url="/teams/$($_.id)/channels`$filter=membershipType eq 'private'"
        }
    }
    $RequestBody = $Request | ConvertTo-Json -Depth 3
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -NotIn @("200","404") | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
        return
    }
    $TeamChannels += $Response.responses | ? status -eq "200"
}
$ChannelsToRemove = @()
$TeamChannels | % {
    $TeamID = $_.id     
    $ChannelsToRemove += $_.body.value | % {                                                                                                                 
        Write-Host "Verwijderen van kanaal $($_.DisplayName) in $(($Teams | ? Id -EQ $TeamID).DisplayName)"
        [PSCustomObject][Ordered]@{
            Id=$_.ID
            Method='DELETE'
            Url="/teams/$TeamID/channels/$($_.id)"
        }
    }
}
$choiceRTN = $host.UI.PromptForChoice("OPKUISEN PRIVEKANALEN", "Opkuisen van $($ChannelsToRemove.Count) kanalen", $options, 1)
if ( $choiceRTN -eq 1 ) {
    return
}

for($i=0;$i -lt $ChannelsToRemove.count;$i+=20){                                                                                                                                              
    Write-Progress -Activity "Privékanalen verwijderen..." -Status "$i/$($ChannelsToRemove.Count) gedaan" -PercentComplete ($i / $ChannelsToRemove.Count * 100)
    $Request = @{}                
    $Request['requests'] = ($ChannelsToRemove[$i..($i+19)])
    $RequestBody = $Request | ConvertTo-Json -Depth 3
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "204" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
}

# TEAMS ARCHIVEREN
$TeamsToArchive = @()
$Teams | % {
    Write-Host "Archiveren van team $($_.DisplayName)"
    $Headers = [PSCustomObject][Ordered]@{"Content-Type"="application/json"}
    $Body = [PSCustomObject][Ordered]@{"shouldSetSpoSiteReadOnlyForMembers"=$true}
    $TeamsToArchive += [PSCustomObject][Ordered]@{
        Id=$_.DisplayName
        Method='POST'
        Url="/teams/$($_.ID)/archive"
        Headers=$Headers
        Body=$Body
    }
}
$choiceRTN = $host.UI.PromptForChoice("ARCHIVEREN KANALEN", "Archiveren van $($TeamsToArchive.Count) teams", $options, 1)
if ( $choiceRTN -eq 1 ) {
    return
}
for($i=0;$i -lt $TeamsToArchive.count;$i+=20){                                                                                                                                              
    Write-Progress -Activity "Teams archiveren..." -Status "$i/$($TeamsToArchive.Count) gedaan" -PercentComplete ($i / $TeamsToArchive.Count * 100)
    $Request = @{}                
    $Request['requests'] = ($TeamsToArchive[$i..($i+19)])
    $RequestBody = $Request | ConvertTo-Json -Depth 3
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "204" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
}

$Groups = Get-MgGroup -Filter "startswith(displayname,'$($Yearcode)') and groupTypes/any(c:c eq 'DynamicMembership')"
# TEAMS ARCHIVEREN
$GroupsToPause = @()
$Groups | % {
    $Headers = [PSCustomObject][Ordered]@{"Content-Type"="application/json"}
    $Body = [PSCustomObject][Ordered]@{"MembershipRuleProcessingState"="Paused"}
    $GroupsToPause += [PSCustomObject][Ordered]@{
        Id=$_.DisplayName
        Method='PATCH'
        Url="/groups/$($_.ID)"
        Headers=$Headers
        Body=$Body
    }
}
for($i=0;$i -lt $GroupsToPause.count;$i+=20){                                                                                                                                              
    Write-Progress -Activity "Lidmaatschap bevriezen..." -Status "$i/$($GroupsToPause.Count) gedaan" -PercentComplete ($i / $GroupsToPause.Count * 100)
    $Request = @{}                
    $Request['requests'] = ($GroupsToPause[$i..($i+19)])
    $RequestBody = $Request | ConvertTo-Json -Depth 3
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "204" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
}

# Oude teams schrappen
$Response = Invoke-GraphRequest -Uri "https://graph.microsoft.com/beta/teams?`$filter=startswith(displayName,'$($OldYearCode)')"
$Teams = $Response.value
while ($Response.'@odata.nextLink') {
    $Response = Invoke-GraphRequest -Uri $Response.'@odata.nextLink'
    $Teams += $Response.value
}

$TeamsToRemove = @()
$Groups | % {
    Write-Host "Verwijderen van team $($_.DisplayName)"
    $TeamsToRemove += [PSCustomObject][Ordered]@{
        Id=$_.DisplayName
        Method='DELETE'
        Url="/groups/$($_.ID)"
    }
}
$choiceRTN = $host.UI.PromptForChoice("VERWIJDEREN TEAMS", "Verwijderen van $($TeamsToRemove.Count) kanalen", $options, 1)
if ( $choiceRTN -eq 1 ) {
    return
}

for($i=0;$i -lt $TeamsToRemove.count;$i+=20){                                                                                                                                              
    Write-Progress -Activity "Oude groepen verwijderen..." -Status "$i/$($TeamsToRemove.Count) ge#>#daan" -PercentComplete ($i / $TeamsToRemove.Count * 100)
    $Request = @{}                
    $Request['requests'] = ($TeamsToRemove[$i..($i+19)])
    $RequestBody = $Request | ConvertTo-Json -Depth 3
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "204" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
}
