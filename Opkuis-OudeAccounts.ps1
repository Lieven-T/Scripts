$YearCode = "2122_"
Connect-Graph -Scopes @("Group.ReadWrite.All")
$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

################################
### OPKUIS LEGE KLASSEN O365 ###
################################

Connect-Graph
$XmlFileLocation = "\\orc-dc1\c$\Program Files\ADWeaver\klassen"
$SmsClassList = @()
$SmsClassList = Get-ChildItem $XmlFileLocation | % { ([xml](Get-Content -Path $_.FullName)).Objs.ChildNodes } | % { $_.'#text'}

$Groups = Get-MgGroup -Filter "startswith(displayName,'romerocollege')" -All
$Groups = $Groups | ? { $_.DisplayName -Match "romerocollege.*_[1-7]" -and ($_.displayName -split "_")[-1] -notin $SmsClassList }
$Teams = Get-MgGroup -Filter "startswith(displayName,'$YearCode')" -All

# ALLE GROEPLEDEN OPHALEN
$GroupMembers = @()
for($i=0;$i -lt $Groups.count;$i+=20){                                                                                                                                              
    Write-Progress -Activity "Groepsleden zoeken..." -Status "$i/$($Groups.Count) gedaan" -PercentComplete ($i / $Groups.Count * 100)
    $Request = @{}
    $Request.requests = $Groups[$i..($i+19)] | % {
        [PSCustomObject][Ordered]@{
            id=$_.DisplayName
            method='GET'
            Url="/groups/$($_.id)/members"
        }
    }
    $RequestBody = $Request | ConvertTo-Json -Depth 3
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "200" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
    $GroupMembers += $Response.responses
}

$GroupsToRemove = @()
$Groups | % {
    if (($GroupMembers | ? id -eq $_.DisplayName).body.value.count) {
        Write-Error "Ongeldige klas met leerlingen: $($_.DisplayName)"
        return
    }
    Write-Host "Opkuisen klas $($_.DisplayName)"
    $GroupsToRemove += [PSCustomObject][Ordered]@{
        Id=$_.DisplayName
        Method='DELETE'
        Url="/groups/$($_.ID)"
    }
}
$choiceRTN = $host.UI.PromptForChoice("OPKUISEN KLASSEN", "Opkuisen van $($GroupsToRemove.Count) klassen", $options, 1)
if ( $choiceRTN -eq 1 ) {
    return
}

for($i=0;$i -lt $GroupsToRemove.count;$i+=20) {
    Write-Progress -Activity "Ongeldige klassen verwijderen..." -Status "$i/$($GroupsToRemove.Count) gedaan" -PercentComplete ($i / $GroupsToRemove.Count * 100)
    $Request = @{}           
    $Request['requests'] = $GroupsToRemove[$i..($i+19)]
    $RequestBody = $Request | ConvertTo-Json -Depth 4
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "204" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
}

$TeamNames = $GroupsToRemove | % {
    "$YearCode$(($_.Id -split '_')[-1])"
}
$Teams = $Teams | ? DisplayName -in $TeamNames
$TeamsToRemove = @()
$Teams | % {
    $TeamsToRemove += [PSCustomObject][Ordered]@{
        Id=$_.DisplayName
        Method='DELETE'
        Url="/groups/$($_.ID)"
    }
}
for($i=0;$i -lt $TeamsToRemove.count;$i+=20){                                                                                                                                              
    Write-Progress -Activity "Ongeldige teams verwijderen..." -Status "$i/$($TeamsToRemove.Count) gedaan" -PercentComplete ($i / $TeamsToRemove.Count * 100)
    $Request = @{}                
    $Request['requests'] = ($TeamsToRemove[$i..($i+19)])
    $RequestBody = $Request | ConvertTo-Json -Depth 3
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "204" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
}

######################
# DEELTIJDS OPKUISEN #
######################

$Groups = Get-MgGroup -Filter "startswith(displayName,'trefpunt')" -All | ? DisplayName -Match "\d"
$GroupMembers = @()
for($i=0;$i -lt $Groups.count;$i+=20){                                                                                                                                              
    Write-Progress -Activity "Groepsleden zoeken..." -Status "$i/$($Groups.Count) gedaan" -PercentComplete ($i / $Groups.Count * 100)
    $Request = @{}
    $Request.requests = $Groups[$i..($i+19)] | % {
        [PSCustomObject][Ordered]@{
            id=$_.DisplayName
            method='GET'
            Url="/groups/$($_.id)/members"
        }
    }
    $RequestBody = $Request | ConvertTo-Json -Depth 3
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "200" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
    $GroupMembers += $Response.responses
}

$GroupsToRemove = @()
$Groups | % {
    if (($GroupMembers | ? id -eq $_.DisplayName).body.value.count) {
        return
    }
    Write-Host "Opkuisen klas $($_.DisplayName)"
    $GroupsToRemove += [PSCustomObject][Ordered]@{
        Id=$_.DisplayName
        Method='DELETE'
        Url="/groups/$($_.ID)"
    }
}
$choiceRTN = $host.UI.PromptForChoice("OPKUISEN KLASSEN", "Opkuisen van $($GroupsToRemove.Count) klassen", $options, 1)
if ( $choiceRTN -eq 1 ) {
    return
}
for($i=0;$i -lt $GroupsToRemove.count;$i+=20) {
    Write-Progress -Activity "Lege CLW-klassen verwijderen..." -Status "$i/$($GroupsToRemove.Count) gedaan" -PercentComplete ($i / $GroupsToRemove.Count * 100)
    $Request = @{}           
    $Request['requests'] = ($GroupsToRemove[$i..($i+19)])
    $RequestBody = $Request | ConvertTo-Json -Depth 4
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "204" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
}

#####################################
### OPKUIS ONBESTAANDE KLASSEN AD ###
#####################################

 $XmlFileLocation = "\\orc-dc1\c$\Program Files\ADWeaver\klassen"
[System.Collections.ArrayList]$SmsClassList = @()
$SmsClassList = Get-ChildItem $XmlFileLocation | % { ([xml](Get-Content -Path $_.FullName)).Objs.ChildNodes } | % { $_.'#text'}
Import-Module ActiveDirectory

Set-Location AD:
Set-Location "DC=campus,DC=romerocollege,DC=be"
Set-Location "OU=leerlingen,OU=gebruikers,OU=school"
Get-ChildItem | ? Name -NE "inactief" | ? Name -NotIn $SmsClassList | % {
    Write-Host $_.Name
    Remove-ADOrganizationalUnit -Identity $_.DistinguishedName -Confirm:$False
}


################################
### VERZAMELEN SPOOKACCOUNTS ###
################################

$SpookAccountID = "864f9916-6ebd-44ed-abc8-9629495142bc"

# Zoek alle leden van adweaver-groepen die géén klassen zijn
$AllGroups = Get-MgGroup -Filter "startswith(displayName,'hetlaar') or startswith(displayName,'basisromero') or startswith(displayName,'romerocollege') or startswith(displayName,'trefpunt')" -all
$Groups = $AllGroups | ? DisplayName -NotMatch "basisromero_\d[A-Z]|cvw_\d|romerocollege_\d|trefpunt_\w{2} \w{2}"
$GroupMembers = @()
for($i=0;$i -lt $Groups.count;$i+=20){                                                                                                                                              
    Write-Progress -Activity "Groepsleden zoeken..." -Status "$i/$($Groups.Count) gedaan" -PercentComplete ($i / $Groups.Count * 100)
    $Request = @{}
    $Request.requests = @()
    $Groups[$i..($i+19)] | % {
        $Request.requests += [PSCustomObject][Ordered]@{
            id=$_.DisplayName
            method='GET'
            Url="/groups/$($_.id)/members"
        }
    }
    $RequestBody = $Request | ConvertTo-Json -Depth 3
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "200" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
    $GroupMembers += $Response.responses
}

$MemberList = $GroupMembers | % { 
    [PSCustomObject][Ordered]@{
        Id=$_.id
        Members=$_.body.value
    }
}
$MemberList | ? { $_.Members.Count -ge 100 } | % {
    Write-Host "Ophalen alle leden van groep $($_.Id)"
    $_.Members = Get-MgGroupMember -GroupId ($Groups | ? DisplayName -eq $_.Id).Id -All
}

$SpookAccounts = Get-MgGroupMember -GroupId $SpookAccountID -All
$RegularUserIDs = $MemberList | % { $_.Members | select -ExpandProperty ID } | Select -Unique
$RegularAndGhostUserIDs = $RegularUserIDs + ($SpookAccounts | select -ExpandProperty ID) | select -Unique
$AllUsers = Get-MgUser -all | ? UserPrincipalName -NotMatch "^package|#EXT#"

# Alle gebruikers die niét in een ADWeaver-hoofdgroep of de spookgroep zitten, toevoegen aan de spookgroep
$UsersToAdd = @()
$AllUsers | ? Id -NotIn $RegularAndGhostUserIDs | % {
    Write-Host "Toevoegen $($_.UserPrincipalName) aan Spookaccounts"
    $Headers = [PSCustomObject][Ordered]@{"Content-Type"="application/json"}
    $Body = [PSCustomObject][Ordered]@{ 
        "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($_.id)"
    }
    $UsersToAdd += [PSCustomObject][Ordered]@{
        Id=$_.UserPrincipalName
        Method='POST'
        Url="groups/$SpookAccountID/members/`$ref"
        Headers=$Headers
        Body=$Body
    }
}
$choiceRTN = $host.UI.PromptForChoice("SPOOKACCOUNTS VERZAMELEN", "Verzamelen van $($UsersToAdd.Count) accounts", $options, 1)
if ( $choiceRTN -eq 1 ) {
    return
}
for($i=0;$i -lt $UsersToAdd.count;$i+=20) {
    Write-Progress -Activity "Accounts verzamelen..." -Status "$i/$($UsersToAdd.Count) gedaan" -PercentComplete ($i / $UsersToAdd.Count * 100)
    $Request = @{}           
    $Request['requests'] = ($UsersToAdd[$i..($i+19)])
    $RequestBody = $Request | ConvertTo-Json -Depth 4
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "204" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
}

#######################################
### OPKUISEN EXTRA EN SPOOKACCOUNTS ###
#######################################

# SPOOKGROEP UITKUISEN
$MembersToRemove = @()
$SpookAccounts | ? ID -in $RegularUserIDs  | % {
    Write-Host "Schrappen $($_.AdditionalProperties.userPrincipalName) uit Spookaccounts"
    $MembersToRemove += [PSCustomObject][Ordered]@{
        Id=$_.AdditionalProperties.userPrincipalName
        Method='DELETE'
        Url="groups/$SpookAccountID/members/$($_.id)/`$ref"
    }
}
$choiceRTN = $host.UI.PromptForChoice("SPOOKACCOUNTS SCHRAPPEN", "Schrappen van $($MembersToRemove.Count) accounts", $options, 1)
if ( $choiceRTN -eq 1 ) {
    return
}
for($i=0;$i -lt $MembersToRemove.count;$i+=20) {
    Write-Progress -Activity "Accounts schrappen..." -Status "$i/$($MembersToRemove.Count) gedaan" -PercentComplete ($i / $MembersToRemove.Count * 100)
    $Request = @{}           
    $Request['requests'] = ($MembersToRemove[$i..($i+19)])
    $RequestBody = $Request | ConvertTo-Json -Depth 4
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "204" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
}

# EXTRA UITKUISEN
$RegularUserWithoutExtraIDs = $MemberList | ? Id -NotMatch "_extra$" | % { $_.Members | select -ExpandProperty ID } | Select -Unique
$MemberList | ? Id -match "_extra$" | % {
    $CurrentGroupName = $_.Id
    $CurrentGroupId = ($Groups | ? DisplayName -eq $CurrentGroupName).Id
    $MembersToRemove = @()
    $_.MemberList | ? ID -in $RegularUserWithoutExtraIDs | % {
        Write-Host "Schrappen $($_.AdditionalProperties.userPrincipalName) uit $CurrentGroupName"
        $MembersToRemove += [PSCustomObject][Ordered]@{
            Id=$_.AdditionalProperties.userPrincipalName
            Method='DELETE'
            Url="groups/$CurrentGroupId/members/$($_.id)/`$ref"
        }
    }
    $choiceRTN = $host.UI.PromptForChoice("EXTRA-ACCOUNTS SCHRAPPEN", "Schrappen van $($MembersToRemove.Count) accounts uit $CurrentGroupName", $options, 1)
    if ( $choiceRTN -eq 1 ) {
        return
    }
    for($i=0;$i -lt $MembersToRemove.count;$i+=20) {
        Write-Progress -Activity "Accounts schrappen..." -Status "$i/$($MembersToRemove.Count) gedaan" -PercentComplete ($i / $MembersToRemove.Count * 100)
        $Request = @{}           
        $Request['requests'] = ($MembersToRemove[$i..($i+19)])
        $RequestBody = $Request | ConvertTo-Json -Depth 4
        $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
        $Response.responses | ? status -ne "204" | % {
            Write-Error "Probleem met $($_.id): $($_.body.error.message)"
        }
    }
}
