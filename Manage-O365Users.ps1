Param
(
    [Parameter(Mandatory=$True,Position=1)]
    [ValidateScript({Test-Path $_})] 
    [string]$InputFile,

    [Parameter(Mandatory=$True,Position=2)]
    [ValidateSet("lk","leraars","ll","leerlingen")] 
    [string]$UserType,

    [Parameter(Position=3)]
    [ValidateSet("schrappen", "aanmaken", "reset-wachtwoord", "filteren")] 
    [string]$Action = "aanmaken"
)

[string]$FirstNameColumn = 'Voornaam'
[string]$LastNameColumn = 'Achternaam'
[string]$EmailColumn = 'Email'
[string]$CreatedUsersFile = 'O365-nieuwe-gebruikers.xlsx'
[string]$DefaultPassword = 'RomeroGeheim!'
[string]$SMS = 'romerocollege'

function Create-Password {
    $Vowels = @('a','e','i','o','u')
    $Consonants = @('b','c','d','f','g','h','k','l','m','n','p','r','s','t','v','v','w','x','z')

    $Password = (Get-Random -InputObject $Consonants).ToUpper()
    $Password += Get-Random -InputObject $Vowels
    $Password += Get-Random -InputObject $Consonants
    $Password += Get-Random -InputObject $Vowels
    $Password += (Get-Random -minimum 0 -maximum 9999).ToString('0000')

    return $Password
}

function Remove-Diacritics {
    Param ([String]$src = [String]::Empty)
    $normalized = $src.Normalize([Text.NormalizationForm]::FormD)
    $sb = new-object Text.StringBuilder
    $normalized.ToCharArray() | % { 
        if( [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($_)
        }

    }
    $sb.ToString()
}

function Create-Email {
    Param ($Account)
    
    if ($Account.Email) {
        return
    } 

    # Compose email address as firstname.lastname@domainname. Diacritics are filtered out and replaced by plain ASCII characters
    if ($_.$LastNameColumn.Length -ne 0) {
	    $Email = "$(Remove-Diacritics $_.$FirstNameColumn.ToLower()).$(Remove-Diacritics $_.$LastNameColumn.ToLower())@$($Domain)"
    } else {
        $Email = "$(Remove-Diacritics $_.$FirstNameColumn.ToLower())@$($Domain)"
    }

	# Strip all spaces, single quotes and dashes
    $Email = $Email -replace '[-\s'']',''

    if ($Account.PSObject.Properties.Match('Email').Count) {
        $Account.Email = $Email
    }
    else
    {
        $Account | Add-Member -Name "Email" -Value $Email -MemberType NoteProperty
    }
}

# Import and verify data
$InputData = $null
if ($InputFile -match '.csv')
{
    $InputData = Import-Csv -Path $InputFile  -Delimiter ";"
}
else {
    $SheetName = (Get-ExcelSheetInfo -Path $InputFile)[0].Name
    $InputData = Import-Excel -Path $InputFile -WorksheetName $SheetName
}

if ($InputData.Length -eq 0)
{
    Write-Error "Bestand `'$InputFile`' bevat geen gegevens"
    Exit
}

# FirstNameColumn and LastNameColumn are required, abort if these aren't found
$ColumnNames = $InputData | Get-Member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name'
if ($ColumnNames -notcontains $FirstNameColumn)
{
    Write-Error "Kolom `'$FirstNameColumn`' zou de voornaam moeten bevatten, maar deze kolom bestaat niet."
    Exit
}

if ($ColumnNames -notcontains $LastNameColumn)
{
    Write-Error "Kolom `'$LastNameColumn`' zou de achternaam moeten bevatten, maar deze kolom bestaat niet."
    Exit
}

if ($UserType -notin @('leerlingen', 'leraars', 'll', 'lk'))
{
    Write-Error "Parameter `$UserType is `'$UserType`' maar kan enkel `'leerlingen`'/`'ll`' of `'leraars`'/`'lk`' zijn"
    Exit
}

$Date = Get-Date
$CreatedUsersFile = '' + $Date.Year + $Date.Month.ToString('00') + $Date.Day.ToString('00') + '-' + $CreatedUsersFile 

if (Test-Path $CreatedUsersFile)
{
    Remove-Item $CreatedUsersFile -Confirm -ErrorAction SilentlyContinue
    if (Test-Path $CreatedUsersFile)
    {
        throw "Bestand $CreatedUsersFile bestaat reeds"
    }
}

# Use these credentials to connect to O365 and import the commands provided by O365.
try {
    Get-AzureADTenantDetail | Out-Null
} catch {
    Connect-AzureAD -ErrorAction Stop
}

$License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
$License.SkuId = "78e66a63-337a-4a9a-8959-41c6654dfb56"
$LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
$LicensesToAssign.AddLicenses = $License
$Domain = 'romerocollege.be'
$GroupId = (Get-AzureADGroup -SearchString "$($SMS)_extra").ObjectId
If ($UserType -in ('leerlingen', 'll'))
{
    $License.SkuId = "e82ae690-a2d5-4d76-8d30-7c6e01e6022e"
    $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
    $LicensesToAssign.AddLicenses = $License
    $Domain = 'student.romerocollege.be'
}

$Cancel = New-Object System.Management.Automation.Host.ChoiceDescription 'Annuleren','Annuleert de huidige bewerking'
$Ok = New-Object System.Management.Automation.Host.ChoiceDescription 'OK','Voert de huidige bewerking uit'
$Options = [System.Management.Automation.Host.ChoiceDescription[]] ($Cancel,$Ok)
 
switch($Action) {
    "aanmaken" {
	    # When there are user to create, run through the list and process each item
        $InputData | % {

            Create-Email $_
        
            $Password = $DefaultPassword
            $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
            $PasswordProfile.Password = $Password
            $PasswordProfile.ForceChangePasswordNextLogin = $True

		    # Create the user
		    Write-Host "Aanmaken nieuwe account `'$($_.Email)`'";
            if ($_.$LastNameColumn.Length -ne 0) {
                New-AzureADUser -GivenName $_.$FirstNameColumn `
                    -Surname $_.$LastNameColumn `
                    -DisplayName "$($_.$FirstNameColumn) $($_.$LastNameColumn)" `
                    -MailNickName $_.Email.Split("@")[0] `
                    -UserPrincipalName $_.Email `
                    -UsageLocation BE `
                    -AccountEnabled $True `
                    -PasswordProfile $PasswordProfile `
                    -ErrorAction Stop `
                    | Out-Null
            } else {
                New-AzureADUser -DisplayName $_.$FirstNameColumn `
                    -MailNickName $_.Email.Split("@")[0] `
                    -UserPrincipalName $_.Email `
                    -UsageLocation BE `
                    -AccountEnabled $True `
                    -PasswordProfile $PasswordProfile `
                    -ErrorAction Stop `
                    | Out-Null
            }

            $User = Get-AzureADUser -ObjectId $_.Email
            Add-AzureADGroupMember -ObjectId $GroupId -RefObjectId $User.ObjectId
            Set-AzureADUserLicense -ObjectId $User.UserPrincipalName -AssignedLicenses $LicensesToAssign
            $_ | Add-Member -Name 'Wachtwoord' -Value $Password -MemberType NoteProperty
        } 
    }

    "reset-wachtwoord" {
        $Choice = $Host.UI.PromptForChoice('Wachtwoorden resetten', 'Wilt u de wachtwoorden resetten voor de lijst van gebruikers?', $options , 1)

        if ($Choice -eq 1) {
            $InputData | % {

                Create-Email $_
        
		        # Reset the password
		        Write-Host "Resetten wachtwoord voor `'$($_.Email)`'"

                $User = Get-AzureADUser -ObjectId $_.Email
                if ($User -ne $null)
                {
                    $Password = $DefaultPassword
                    Set-AzureADUserPassword -ForceChangePasswordNextLogin $True -Password (ConvertTo-SecureString $Password -AsPlainText -Force) -ObjectId $_.Email
                    $_ | Add-Member -Name 'Wachtwoord' -Value $Password -MemberType NoteProperty
                }
            } 
        }
    }

    "schrappen" {
        $Choice = $Host.UI.PromptForChoice('Gebruikers verwijderen', 'Wilt u de lijst van gebruikers verwijderen?', $options , 1)

        if ($Choice -eq 1) {
            $InputData | ForEach-Object {
                Create-Email $_

		        # Remove the user
                Write-Host "Schrappen account `'$($_.Email)`'"
                Remove-AzureADUser -ObjectId $_.Email
            }
        }
    }

    "filteren" {
        if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue))
        {
            $LiveCred = Get-Credential
            $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
            Import-PSSession $Session
        }
        $InputData | % {
            $LastLogon = $null
            Create-Email $_
            if (Get-Mailbox $_.Email)
            {
                $LastLogon = (Get-MailboxStatistics $_.Email | Select LastLogonTime).LastLogonTime
                if (-not $LastLogon) { $LastLogon = "Niet aangemeld" }
                $_ | Add-Member -Name "LastLogon" -Value $LastLogon -MemberType NoteProperty
            } else {
                $_ | Add-Member -Name "LastLogon" -Value "Onbestaand" -MemberType NoteProperty
            }
        }
    }
}
$InputData | Export-Excel $CreatedUsersFile
