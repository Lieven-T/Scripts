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

$Users = $InputData | Select -Property $ClassColumn,$TeacherColumn,$RoleColumn -Unique
$Users | % {
    $_.$ClassColumn = $_.$ClassColumn -replace "\d{4}_",""
    $_.$TeacherColumn = $(($_.$TeacherColumn -split "@")[0])
}
$ClassCodes = $Users | Select -ExpandProperty $ClassColumn -Unique | % { "$YearCode$_" }

$Classes = for($i = 1;$i -lt 8; $i++) { 
    Get-AzureADGroup -Filter "startswith(displayname,'$($YearCode)$i')" -All $true
}

$Classes | ? { ($_.DisplayName -in $ClassCodes } | % {
    $Class = $_
    $Users | ? { 
        ($_.$ClassColumn -eq ($_.DisplayName -replace $YearCode,"")) -and ($_.$TeacherColumn -notin (Get-AzureADGroupOwner -ObjectId $Class.ObjectId | select -ExpandProperty UserPrincipalName | % { ($_ -split "@")[0] } ) )
    } | % {
        Write-Host "Toevoegen van $($_.$TeacherColumn) aan klas $($_.$ClassColumn)"
        $User = Get-AzureADUser -Filter "UserPrincipalName eq '$($_.$TeacherColumn)@romerocollege.be'"
        Add-AzureADGroupOwner -ObjectId $Class.ObjectId -RefObjectId $User.ObjectId
    }
}

# ADWeaver schrappen als eigenaar van klas
$User = Get-AzureADUser -Filter "UserPrincipalName eq 'adweaver@romerocollege.be'"
$Classes | ? { ($_.DisplayName -split "_")[1] -in $ClassCodes }| % {
    Write-Host "ADWeaver schrappen uit klas $($_.DisplayName)"
    Remove-AzureADGroupOwner -ObjectId $_.ObjectId -OwnerId $User.ObjectId
}
