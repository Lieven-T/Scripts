Param
(
    [Parameter(Mandatory=$True,Position=1)]
    [ValidateScript({Test-Path $_})] 
    [string]$InputFile
)

# FORMAAT: ZONDER @ROMEROCOLLEGE, VOLLEDIGE NAAM VAN TEAM

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

[string]$YearCode = "2122_KR"

$Users = $InputData | Select -Property $ClassColumn,$TeacherColumn,$RoleColumn -Unique
$ClassCodes = $Users | Select -Unique -ExpandProperty $ClassColumn

# Klassenraden aanmaken
$Klassenraden = Get-AzureADGroup -SearchString $YearCode -All $true
$ClassCodes | ? { $_ -notin ($Klassenraden | select -ExpandProperty DisplayName)} | % {
    $DisplayName = $_
    Write-Host "Aanmaken klassenraad: $DisplayName"
    if (Get-AzureADGroup -Filter "displayname eq '$DisplayName'") {
        Write-Host "    Team bestaat al: "
        Get-AzureADGroup -Filter "displayname eq '$DisplayName'"
    } else {
        $Klassenraad = New-Team -DisplayName $DisplayName -MailNickName ($DisplayName -replace " ","_")
    }
}

$Klassenraden = Get-AzureADGroup -SearchString $YearCode -all $True
$Users | % {
    # Leraar toevoegen aan klassenraad
    Write-Host "Toevoegen van $($_.$TeacherColumn) aan klassenraad $($_.$ClassColumn) als $($_.$RoleColumn)"
    $Klas = $Klassenraden | ? DisplayName -eq $_.$ClassColumn 
    $User = Get-AzureADUser -SearchString $_.$TeacherColumn
    Write-Host "    $($User.UserPrincipalName) -> $($Klas.DisplayName)"
    if ($_.$RoleColumn -eq "eigenaar" ) {
        Write-Host "    Eigenaar"
        Add-AzureADGroupOwner -ObjectId $Klas.ObjectId -RefObjectId $User.ObjectId
    } else {
        Write-Host "    Lid"
        Add-AzureADGroupMember -ObjectId $Klas.ObjectId -RefObjectId $User.ObjectId
    }
}

# Mezelf schrappen als eigenaar van klassenraad
$User = get-azureaduser -searchstring "lieven.tronckoe"
$Klassenraden | % {
    Write-Host "Mezelf schrappen uit klassenraad $($_.DisplayName)"
    Remove-AzureADGroupOwner -ObjectId $_.ObjectId -OwnerId $User.ObjectId
}