$YearCode = "2122_"

################################
### OPKUIS LEGE KLASSEN O365 ###
################################

# TODO: batching/herwerken
Connect-Graph
$XmlFileLocation = "\\orc-dc1\c$\Program Files\ADWeaver\klassen"
$SmsClassList = @()
$SmsClassList = Get-ChildItem $XmlFileLocation | % { ([xml](Get-Content -Path $_.FullName)).Objs.ChildNodes } | % { $_.'#text'}

$Groups = Get-MgGroup -Filter "startswith(displayName,'romerocollege')" -All
$Groups = $Groups | ? { $_.DisplayName -Match "romerocollege.*_[1-7]" -and ($_.displayName -split "_")[1] -notin $SmsClassList }

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
    $Request['requests'] = ($GroupsToRemove[$i..($i+19)])
    $RequestBody = $Request | ConvertTo-Json -Depth 4
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

Write-Host "Trefpunt"
$trefpuntLeraars = Get-AzureADGroup -SearchString trefpunt_ -All $true | ? DisplayName -NotMatch "_\w{2} \w{2}|_extra|_disabled|_leerlingen" | % { Get-AzureADGroupMember -ObjectId $_.objectid -All $true } | Select -ExpandProperty ObjectId -Unique
Get-AzureADGroup -SearchString trefpunt_extra -All $true | % { Get-AzureADGroupMember -ObjectId $_.objectid -All $true } | ? { $_.UserPrincipalName -Match "@romero" -and $_.ObjectId -in $trefpuntLeraars }

Write-Host "Romero"
$romeroLeraars = Get-AzureADGroup -SearchString romerocollege_ -All $true | ? DisplayName -notmatch "cvw_\d|romerocollege_\d|_extra|_disabled|_leerlingen|_OKAN" | % { Get-AzureADGroupMember -ObjectId $_.objectid -All $true } | Select -ExpandProperty ObjectId -Unique
Get-AzureADGroup -SearchString romerocollege_extra -All $true | % { Get-AzureADGroupMember -ObjectId $_.objectid -All $true } | ? { $_.UserPrincipalName -Match "@romero" -and $_.ObjectId -in $romeroLeraars }

Write-Host "Basis"
$basisLeraars = Get-AzureADGroup -SearchString basisromero_ -All $true | ? DisplayName -notmatch "_extra|_disabled|_leerlingen|_\d[A-Z]" | % { Get-AzureADGroupMember -ObjectId $_.objectid -All $true } | Select -ExpandProperty ObjectId -Unique
Get-AzureADGroup -SearchString basisromero_extra -All $true | % { Get-AzureADGroupMember -ObjectId $_.objectid -All $true } | ? { $_.UserPrincipalName -Match "@romero" -and $_.ObjectId -in $basisLeraars }


#######################################
### OPKUIS ONBEHEERDE ACCOUNTS O365 ###
#######################################

Connect-AzureAD

$OudeLeraars = (Get-AzureADGroup -SearchString oudeleraars).ObjectId
$OudeLln = (Get-AzureADGroup -SearchString oudelln).ObjectId

Write-Host "Het Laar"
$hetlaarusers = Get-AzureADGroup -SearchString hetlaar -All $true | % { Get-AzureADGroupMember -ObjectId $_.objectid -All $true } | Select -ExpandProperty ObjectId -Unique;

Write-Host "Trefpunt"
$trefpuntusers = Get-AzureADGroup -SearchString trefpunt_ -All $true | % { Get-AzureADGroupMember -ObjectId $_.objectid -All $true } | Select -ExpandProperty ObjectId -Unique;

Write-Host "Romero"
$romerousers = Get-AzureADGroup -SearchString romerocollege_ -All $true | % { Get-AzureADGroupMember -ObjectId $_.objectid -All $true } | Select -ExpandProperty ObjectId -Unique;

Write-Host "Basis"
$basisusers = Get-AzureADGroup -SearchString basisromero_ -All $true | % { Get-AzureADGroupMember -ObjectId $_.objectid -All $true } | Select -ExpandProperty ObjectId -Unique;

$allusers = ($hetlaar + $trefpuntusers + $romerousers + $basisusers) | Select -Unique

Write-Host "Alles"
$users = Get-AzureADUser -All $true

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


#Get-AzureADGroupMember -ObjectId $OudeLeraars | ? { ($_[0].Mail).Split("@")[0] -notin $geldig } | % {
#    Write-Host $_.Mail
    #Set-AzureADUser -ObjectId $_.ObjectId -AccountEnabled $false
#}


# $romero_extra = (Get-AzureADGroup -SearchString romerocollege_extra).ObjectId

#Get-AzureADGroupMember -ObjectId $OudeLeraars | ? { ($_[0].Mail).Split("@")[0] -in $geldig } | % {
#    Write-Host $_.Mail
#    Add-AzureADGroupMember -ObjectId $romero_extra -RefObjectId $_.ObjectId
#    Remove-AzureADGroupMember -ObjectId $OudeLeraars -MemberId $_.ObjectId
#}

