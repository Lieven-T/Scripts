Param
(
    [Parameter(Mandatory=$True,Position=1)]
    [ValidateScript({Test-Path $_})] 
    [string]$InputFile
)

[string]$ClassColumn = 'Klas'
[string]$TeacherColumn = 'Leraar'
[string]$RoleColumn = 'Rol'

Write-Error "DEBUGGEN!"
return

$yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes"
$no = New-Object System.Management.Automation.Host.ChoiceDescription "&No"
$options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)

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

[string]$YearCode = "2122_KR_"

$InputData = $InputData | Select -Property $ClassColumn,$TeacherColumn -Unique
$InputData | % {
    $_.$ClassColumn = $_.$ClassColumn -match $YearCode ? $_.$ClassColumn : "$YearCode$($_.$ClassColumn)"
    $_.$TeacherColumn = $_.$TeacherColumn -match "@" ? $_.$TeacherColumn : "$($_.$TeacherColumn)@romerocollege.be"
}
$InputClasses = $InputData | Select -ExpandProperty $ClassColumn -Unique
$InputUsers = $InputData | Select -ExpandProperty $TeacherColumn -Unique

Connect-Graph -Scopes @("Group.ReadWrite.All")

# NODIGE GROEPEN EN USERS OPHALEN UIT AZUREAD
$AzureGroups = Get-MgGroup -Filter "startswith(displayName,'$YearCode')" -All
$AzureUsers = Get-MgUser -All | ? UserPrincipalName -in $InputUsers
$AzureUPNs = $AzureUsers | select -ExpandProperty UserPrincipalName
$InputUsers | ? { $_ -notin $AzureUPNs} | %{
    Write-Error "Gebruiker niet gevonden: $_"
}
$InputData = $InputData | ? { $_.$TeacherColumn -in $AzureUPNs }

$InputClasses | ? { $_.$ClassColumn -notin $AzureGroups.DisplayName } | % {
    # TODO: AANMAKEN
    Write-Host "Aanmaken $_"
}

$TeamsToCreate = @()
$ClassCodes | ? { $_ -notin $Klassenraden.DisplayName } | % {
    Write-Host "Aanmaken van team $_"
    $Headers = [PSCustomObject][Ordered]@{"Content-Type"="application/json"}
    $Body = [PSCustomObject][Ordered]@{
        "template@odata.bind"="https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
        displayName=$_
        description=$_
        mailNickName=$_
    }
    $TeamsToCreate += [PSCustomObject][Ordered]@{
        Id=$_
        Method='POST'
        Url="/teams"
        Headers=$Headers
        Body=$Body
    }
}
$choiceRTN = $host.UI.PromptForChoice("AAMAKEN TEAMS", "Aanmaken van $($TeamsToCreate.Count)", $options, 1)
if ( $choiceRTN -eq 1 ) {
    return
}

$InitialResponses = @()
for($i=0;$i -lt $TeamsToCreate.count;$i+=20) {
    Write-Progress -Activity "Teams aanmaken..." -Status "$i/$($TeamsToCreate.Count) gedaan" -PercentComplete ($i / $TeamsToCreate.Count * 100)
    $Request = @{}           
    $Request['requests'] = ($TeamsToCreate[$i..($i+19)])
    $RequestBody = $Request | ConvertTo-Json -Depth 4
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "202" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
    $InitialResponses += $Response.responses
}

$Actions = @()
$InitialResponses | % {
    $Actions += @([PSCustomObject][Ordered]@{
        Id=$_.id
        Method='GET'
        Url=$_.headers.Location
    })
}

while($Actions.count) {
    Write-Host "Wachten op $($Actions.count)..."
    Start-Sleep 45
    $Responses = @()
    for($i=0;$i -lt $Actions.count;$i+=20) {
        Write-Progress -Activity "Opvragen van $($Actions.Count) acties" -Status "$i/$($Actions.Count) gedaan" -PercentComplete ($i / $Actions.Count * 100)
        $Request = @{}           
        $Request['requests'] = ($Actions[$i..($i+19)])
        $RequestBody = $Request | ConvertTo-Json -Depth 4
        $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
        $Response.responses | ? status -ne "200" | % {
            Write-Error "Probleem met $($_.id): $($_.body.error.message)"
        }
        $Responses += $Response.responses | ? { $_.status -eq "200" -and $Response.responses[0].body.status -ne "succeeded" }
    }
    $Actions = $Actions | ? Id -in $Responses.Id
}

$AzureGroups = Get-MgGroup -Filter "startswith(displayName,'$YearCode')" -All

# ALLE GROEPEIGENAARS OPHALEN
$AzureOwners = @()
for($i=0;$i -lt $AzureGroups.count;$i+=20){                                                                                                                                              
    Write-Progress -Activity "Groepseigenaars zoeken..." -Status "$i/$($AzureGroups.Count) gedaan" -PercentComplete ($i / $AzureGroups.Count * 100)
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

# ALLE GROEPLEDEN OPHALEN
$AzureMembers = @()
for($i=0;$i -lt $AzureGroups.count;$i+=20){                                                                                                                                              
    Write-Progress -Activity "Groepsleden zoeken..." -Status "$i/$($AzureGroups.Count) gedaan" -PercentComplete ($i / $AzureGroups.Count * 100)
    $Request = @{}
    $Request.requests = $AzureGroups[$i..($i+19)] | % {
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
    $AzureMembers += $Response.responses
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
    $InputGroupOwners = $InputData | ? { $_.$ClassColumn -eq $ClassCode -and $_.$RoleColumn -eq 'eigenaar' } | select -ExpandProperty $TeacherColumn
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
    Write-Progress -Activity "Eigenaars toevoegen..." -Status "$i/$($OwnersToAdd.Count) gedaan" -PercentComplete ($i / $OwnersToAdd.Count * 100)
    $Request = @{}           
    $Request['requests'] = ($OwnersToAdd[$i..($i+19)])
    $RequestBody = $Request | ConvertTo-Json -Depth 4
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "204" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
}

$MembersToAdd = @()
$InputClasses | % {
    $AzureGroup = $AzureGroups | ? DisplayName -eq $_
    if (-not $AzureGroup) {
        Write-Error "Klas niet gevonden: $_"
        return
    }
    $CurrentGroupMembers = $AzureMembers | ? id -eq $_
    $ClassCode = $_
    $InputGroupMembers = $InputData | ? { $_.$ClassColumn -eq $ClassCode -and $_.$RoleColumn -eq 'lid' } | select -ExpandProperty $TeacherColumn
    $GroupMembersToAdd = @()
    $InputGroupMembers | % {
        $AzureUserToAdd = $AzureUsers | ? UserPrincipalName -EQ $_
        if (-not ($CurrentGroupMembers | ? UserPrincipalName -EQ $_)) {
             $GroupMembersToAdd += "https://graph.microsoft.com/v1.0/directoryObjects/$($AzureUserToAdd.id)"
        }
    }
    $Headers = [PSCustomObject][Ordered]@{"Content-Type"="application/json"}
    $Body = [PSCustomObject][Ordered]@{ "members@odata.bind" = $InputGroupMembers }
    $MemberToAdd = [PSCustomObject][Ordered]@{
        id=$_
        method='PATCH'
        Url='/groups/' + $($AzureGroup.id)
        Headers=$Headers
        Body=$Body
    }
    $MembersToAdd += $MemberToAdd
}

for($i=0;$i -lt $MembersToAdd.count;$i+=20) {
    Write-Progress -Activity "Leden toevoegen..." -Status "$i/$($MembersToAdd.Count) gedaan" -PercentComplete ($i / $MembersToAdd.Count * 100)
    $Request = @{}           
    $Request['requests'] = ($MembersToAdd[$i..($i+19)])
    $RequestBody = $Request | ConvertTo-Json -Depth 4
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "204" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
}

$OwnersToRemove = @()
$MezelfUPN = "lieven.tronckoe@romerocollege.be"
$MezelfId = $AzureUsers | ? UserPrincipalName -eq $MezelfUPN
$InputClasses | % {
    $AzureGroup = $AzureGroups | ? DisplayName -eq $_
    if (-not $AzureGroup) {
        Write-Error "Klas niet gevonden: $_"
        return
    }
    $CurrentGroupOwners = $AzureOwners | ? id -eq $_
    if ($CurrentGroupOwners | ? UserPrincipalName -EQ "lieven.tronckoe@romerocollege.be") {
        $OwnerToRemove = [PSCustomObject][Ordered]@{
            id=$_
            method='DELETE'
            Url="/groups/$($AzureGroup.id)/owners/$MezelfId"
        }
        $OwnersToRemove += $OwnerToRemove
    }
}
for($i=0;$i -lt $OwnersToRemove.count;$i+=20) {
    Write-Progress -Activity "Mezelf schrappen..." -Status "$i/$($OwnersToAdd.Count) gedaan" -PercentComplete ($i / $OwnersToAdd.Count * 100)
    $Request = @{}           
    $Request['requests'] = ($OwnersToRemove[$i..($i+19)])
    $RequestBody = $Request | ConvertTo-Json -Depth 4
    $Response = Invoke-GraphRequest -Uri 'https://graph.microsoft.com/beta/$batch' -Body $RequestBody -Method POST -ContentType "application/json"
    $Response.responses | ? status -ne "204" | % {
        Write-Error "Probleem met $($_.id): $($_.body.error.message)"
    }
}
