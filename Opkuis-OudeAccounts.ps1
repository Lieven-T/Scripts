$YearCode = "2122_"

#######################################
### OPKUIS ONBESTAANDE KLASSEN O365 ###
#######################################

Connect-AzureAD
Connect-MicrosoftTeams
$XmlFileLocation = "\\orc-dc1\c$\Program Files\ADWeaver\klassen"
[System.Collections.ArrayList]$SmsClassList = @()
$SmsClassList = Get-ChildItem $XmlFileLocation | % { ([xml](Get-Content -Path $_.FullName)).Objs.ChildNodes } | % { $_.'#text'}
$GroupList = for($i = 1;$i -lt 8; $i++) { Get-AzureADGroup -Filter "startswith(displayname,'romerocollege_$i') or startswith(displayname,'romerocollege_cvw_$i')" -All $true }

$GroupList | % {
    $GroupName = $_.DisplayName 
    $GroupId = $_.ObjectId
    $ClassName = ($GroupName -split "_")[-1]
    if ($ClassName -notin $SmsClassList) {
        Write-Host "Foute klas: $ClassName"
        $Lln = Get-AzureADGroupMember -ObjectId $_.ObjectId
        if ($Lln) {
            Write-Host "    ...heeft leerlingen"
            $Lln | % { Remove-AzureADUser -ObjectId $_.ObjectId }
        } else {
            Get-Team -DisplayName "$YearCode$ClassName" | Remove-Team
            Remove-AzureADGroup -ObjectId $GroupId
        }
    }
}


 Get-AzureADGroup -SearchString "trefpunt_" | ? DisplayName -Match "\d" | ? { -not (Get-AzureADGroupMember -ObjectId $_.ObjectId )} | % { Remove-AzureADGroup -ObjectId $_.ObjectId }


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

