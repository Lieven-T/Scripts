Param
(
    [Parameter(Mandatory=$True,Position=1)]
    [ValidateScript({Test-Path $_})] 
    [string]$InputFile
)

[string]$ClassColumn = 'Klas'
[string]$TeacherColumn = 'Leraar'

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


try {
    Get-AzureADTenantDetail | Out-Null
} catch {
    Connect-AzureAD
}

[string]$YearCode = "2122_"

$Users = $InputData | Select -Property $ClassColumn,$TeacherColumn -Unique
$ClassCodes = $Users | Select -Unique -ExpandProperty $ClassColumn

# Klassen ophalen
$Classes = for($i = 1;$i -lt 8; $i++) { 
    Get-AzureADGroup -Filter "startswith(displayname,'$($YearCode)$i')" -All $true
}

$Users | % {
    # Leraar toevoegen aan klas
    Write-Host "Toevoegen van $($_.$TeacherColumn) aan klas $($_.$ClassColumn)"
    $Class = $Classes | ? DisplayName -EQ "$YearCode$($_.$ClassColumn)"
    $User = Get-AzureADUser -Filter "UserPrincipalName eq '$($_.$TeacherColumn)'"
    Add-AzureADGroupOwner -ObjectId $Class.ObjectId -RefObjectId $User.ObjectId
}

# ADWeaver schrappen als eigenaar van klas
$User = Get-AzureADUser -Filter "UserPrincipalName eq 'adweaver@romerocollege.be'"
$Classes | ? { ($_.DisplayName -split "_")[1] -in $ClassCodes }| % {
    Write-Host "ADWeaver schrappen uit klas $($_.DisplayName)"
    Remove-AzureADGroupOwner -ObjectId $_.ObjectId -OwnerId $User.ObjectId
}
