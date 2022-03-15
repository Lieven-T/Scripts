$YearCode = "2122_"
Connect-Graph -Scopes @("Group.ReadWrite.All")

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
    $GroupsToRemove += [PSCustomObject][Ordered]@{
        Id=$_.DisplayName
        Method='DELETE'
        Url="/groups/$($_.ID)"
    }
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
    Write-Progress -Activity "Ongeldige teams verwijderen..." -Status "$i/$($TeamsToRemove.Count) ge#>#daan" -PercentComplete ($i / $TeamsToRemove.Count * 100)
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

$Groups = Get-MgGroup -Filter "startswith(displayName,'trefpunt')" | ? DisplayName -Match "\d"
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
        return
    }
    $GroupsToRemove += [PSCustomObject][Ordered]@{
        Id=$_.DisplayName
        Method='DELETE'
        Url="/groups/$($_.ID)"
    }
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


##########################
### ONTERECHT IN EXTRA ###
##########################

$Groups = Get-MgGroup -Filter "startswith(displayName,'hetlaar') or startswith(displayName,'basisromero') or startswith(displayName,'romerocollege') or startswith(displayName,'trefpunt')" -all
#$Groups = $Groups | ? DisplayName -NotMatch "basisromero_\d[A-Z]|cvw_\d|romerocollege_\d|trefpunt_\w{2} \w{2}|_extra|_disabled|_leerlingen"
# ALLE GROEPLEDEN OPHALEN: MOET ANDERS
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
$RegularUserIDs = ($GroupMembers | ? Id -notmatch "_extra").body.value | select -Unique Id

$ExtraMembersToRemove = @()
$ExtraGroupId = ($Groups | ? DisplayName -Match "trefpunt_extra").Id
($GroupMembers | ? Id -match "trefpunt_extra").body.value | select -Unique Id | ? { $_ -in $RegularUserIDs } | % {
    $ExtraMembersToRemove += [PSCustomObject][Ordered]@{
        Id=$_.Id
        Method='DELETE'
        Url="/groups/$ExtraGroupId/members/$($_.Id)/`$ref"
    }
}
$ExtraGroupId = ($Groups | ? DisplayName -Match "basisromero_extra").Id
($GroupMembers | ? Id -match "basisromero_extra").body.value | select -Unique Id | ? { $_ -in $RegularUserIDs } | % {
    $ExtraMembersToRemove += [PSCustomObject][Ordered]@{
        Id=$_.Id
        Method='DELETE'
        Url="/groups/$ExtraGroupId/members/$($_.Id)/`$ref"
    }
}
$ExtraGroupId = ($Groups | ? DisplayName -Match "romerocollege_extra").Id
($GroupMembers | ? Id -match "romerocollege_extra").body.value | select -Unique Id | ? { $_ -in $RegularUserIDs } | % {
    $ExtraMembersToRemove += [PSCustomObject][Ordered]@{
        Id=$_.Id
        Method='DELETE'
        Url="/groups/$ExtraGroupId/members/$($_.Id)/`$ref"
    }
}
for($i=0;$i -lt $ExtraMembersToRemove.count;$i+=20) {
    Write-Progress -Activity "Onterecht in extra opkuisen..." -Status "$i/$($ExtraMembersToRemove.Count) gedaan" -PercentComplete ($i / $ExtraMembersToRemove.Count * 100)
    $Request = @{}           
    $Request['requests'] = ($ExtraMembersToRemove[$i..($i+19)])
    $RequestBody = $Request | ConvertTo-Json -Depth 4
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "204" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
}

#######################################
### OPKUIS ONBEHEERDE ACCOUNTS O365 ###
#######################################

$AllUsers = Get-MgUser -all

$OudeLeraars = (Get-AzureADGroup -SearchString oudeleraars).ObjectId
$OudeLln = (Get-AzureADGroup -SearchString oudelln).ObjectId

# Opkuis leraars
$users | ? { $_.ObjectId -notin $allusers -and $_.Mail -match "@romerocollege.be" } | % {
    Write-Host $_.UserPrincipalName
    Add-AzureADGroupMember -ObjectId $OudeLeraars -RefObjectId $_.ObjectId
}

# Opkuis lln
$users | Where-Object { $_.ObjectId -notin $allusers -and $_.Mail -match "@student.romerocollege.be" } | % {
    Write-Host $_.UserPrincipalName
    Add-AzureADGroupMember -ObjectId $OudeLln -RefObjectId $_.ObjectId
}
