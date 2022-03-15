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

[string]$YearCode = "2122_"

$InputData = $InputData | Select -Property $ClassColumn,$TeacherColumn -Unique
$InputData | % {
    $_.$ClassColumn = $_.$ClassColumn -match $YearCode ? $_.$ClassColumn : "$YearCode$($_.$ClassColumn)"
    $_.$TeacherColumn = $_.$TeacherColumn -match "@" ? $_.$TeacherColumn : "$($_.$TeacherColumn)@romerocollege.be"
}
$InputClasses = $InputData | Select -ExpandProperty $ClassColumn -Unique
$InputUsers = $InputData | Select -ExpandProperty $TeacherColumn -Unique

Connect-Graph -Scopes @("Group.ReadWrite.All")

# NODIGE GROEPEN EN USERS OPHALEN UIT AZUREAD
$AzureGroups = Get-MgGroup -Filter "startswith(displayName,'$YearCode')" -All | ? DisplayName -in $InputClasses
$AzureUsers = Get-MgUser -All | ? UserPrincipalName -in $InputUsers
$AzureUPNs = $AzureUsers | select -ExpandProperty UserPrincipalName
$InputUsers | ? { $_ -notin $AzureUPNs} | %{
    Write-Error "Gebruiker niet gevonden: $_"
}
$InputData = $InputData | ? { $_.$TeacherColumn -in $AzureUPNs }

# ALLE GROEPEIGENAARS OPHALEN
$AzureOwners = @()
for($i=0;$i -lt $AzureGroups.count;$i+=20){                                                                                                                                              
    Write-Progress -Activity "Groepsleden zoeken..." -Status "$i/$($AzureGroups.Count) gedaan" -PercentComplete ($i / $AzureGroups.Count * 100)
    $Request = @{}
    $Request.requests = $AzureGroups[$i..($i+19)] | % {
        [PSCustomObject][Ordered]@{
            id=$_.DisplayName
            method='GET'
            Url="/groups/$($_.id)/owners"
        }
    }
    $RequestBody = $Request | ConvertTo-Json -Depth 3
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "200" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
    $AzureOwners += $Response.responses
}

$OwnersToAdd = @()
$InputClasses | % {
    $AzureGroup = $AzureGroups | ? DisplayName -eq $_
    if (-not $AzureGroup) {
        Write-Error "Klas niet gevonden: $_"
        return
    }
    $CurrentGroupOwners = $AzureOwners | ? id -eq $_
    $ClassCode = $_
    $InputGroupOwners = $InputData | ? { $_.$ClassColumn -eq $ClassCode } | select -ExpandProperty $TeacherColumn
    $InputGroupOwners | % {
        $AzureUserToAdd = $AzureUsers | ? UserPrincipalName -EQ $_
        if (-not ($CurrentGroupOwners | ? UserPrincipalName -EQ $_)) {
            $Headers = [PSCustomObject][Ordered]@{"Content-Type"="application/json"}
            $Body = [PSCustomObject][Ordered]@{ 
                "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($AzureUserToAdd.id)"
            }
            $OwnerToAdd = [PSCustomObject][Ordered]@{
                id=$_
                method='POST'
                Url='/groups/' + $($AzureGroup.id)+ '/owners/$ref'
                Headers=$Headers
                Body=$Body
            }
            $OwnersToAdd += $OwnerToAdd
        }
    }
}

for($i=0;$i -lt $OwnersToAdd.count;$i+=20) {
    Write-Progress -Activity "Leerkrachten toevoegen..." -Status "$i/$($OwnersToAdd.Count) gedaan" -PercentComplete ($i / $OwnersToAdd.Count * 100)
    $Request = @{}           
    $Request['requests'] = ($OwnersToAdd[$i..($i+19)])
    $RequestBody = $Request | ConvertTo-Json -Depth 4
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "204" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
}

