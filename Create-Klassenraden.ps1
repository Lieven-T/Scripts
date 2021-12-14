Param
(
    [Parameter(Mandatory=$True,Position=1)]
    [ValidateScript({Test-Path $_})] 
    [string]$InputFile
)

[string]$ClassColumn = 'Klas'
[string]$TeacherColumn = 'Leraar'
[string]$RoleColumn = 'Rol'

$InputData = Import-Excel -Path $InputFile

$ColumnNames = $InputData | Get-Member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name'
if ($ColumnNames -notcontains $ClassColumn)
{
    Write-Error "Kolom `'$ClassColumn`' zou de klas moeten bevatten, maar deze kolom bestaat niet."
    Exit
}
if ($ColumnNames -notcontains $TeacherColumn)
{
    Write-Error "Kolom `'$RoleColumn`' zou de leraar moeten bevatten, maar deze kolom bestaat niet."
    Exit
}
if ($ColumnNames -notcontains $RoleColumn)
{
    Write-Error "Kolom `'$RoleColumn`' zou de rol moeten bevatten, maar deze kolom bestaat niet."
    Exit
}



try {
    Get-AzureADTenantDetail | Out-Null
} catch {
    Connect-AzureAD -ErrorAction Stop
    Connect-MicrosoftTeams
}

[string]$YearCode = "2122_KR_"

$Users = $InputData | Select -Property $ClassColumn,$TeacherColumn,$RoleColumn -Unique
$Users | % {
    $_.$ClassColumn = $_.$ClassColumn -replace "\d{4}_KR_",""
    $_.$TeacherColumn = $(($_.$TeacherColumn -split "@")[0])
}
$ClassCodes = $Users | Select -Unique -ExpandProperty $ClassColumn

# Klassenraden aanmaken
$Klassenraden = Get-AzureADGroup -SearchString $YearCode -All $true
$TeamCodeList = ($Klassenraden | select -ExpandProperty DisplayName) | % { $_ -replace $YearCode,""}
$ClassCodes | ? { $_ -notin $TeamCodeList } | % {
    $DisplayName = "$YearCode$_"
    Write-Host "Aanmaken klassenraad: $DisplayName"
    if (Get-AzureADGroup -Filter "displayname eq '$DisplayName'") {
        Write-Host "    Team bestaat al: "
        Get-AzureADGroup -Filter "displayname eq '$DisplayName'"
    } else {
        New-Team -DisplayName $DisplayName -MailNickName ($DisplayName -replace " ","_")
    }
}

Get-AzureADGroup -SearchString $YearCode -all $True | % {
    $Klas = $_
    Write-Host "Verwerken klassenraad $($Klas.DisplayName)"
    $CurrentMembers = @() + (Get-AzureADGroupMember -ObjectId $klas.ObjectId) + (Get-AzureADGroupOwner -ObjectId $klas.ObjectId) | select -ExpandProperty UserPrincipalName | % { ($_ -split "@")[0] }
    $Users | ? { $_.$ClassColumn -eq ($Klas.DisplayName -replace $YearCode,"") -and ($_.$TeacherColumn -notin $CurrentMembers) } | % {
        Write-Host "    Toevoegen van $($_.$TeacherColumn) aan klassenraad $($_.$ClassColumn) als $($_.$RoleColumn)"
        $User = Get-AzureADUser -Filter "UserPrincipalName eq '$($_.$TeacherColumn)@romerocollege.be'"
        if ($_.$RoleColumn -eq "eigenaar" ) {
            Add-AzureADGroupOwner -ObjectId $Klas.ObjectId -RefObjectId $User.ObjectId
        } else {
            Write-Host "    Lid"
            Add-AzureADGroupMember -ObjectId $Klas.ObjectId -RefObjectId $User.ObjectId
        }
    }
}

# Mezelf schrappen als eigenaar van klassenraad
$User = get-azureaduser -searchstring "lieven.tronckoe"
$Klassenraden | % {
    Write-Host "Mezelf schrappen uit klassenraad $($_.DisplayName)"
    Remove-AzureADGroupOwner -ObjectId $_.ObjectId -OwnerId $User.ObjectId
}