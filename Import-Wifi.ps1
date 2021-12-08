Param
(
    [Parameter(Mandatory=$True,Position=1)]
    [string]$InputFile
)

[string]$HostNameColumn = 'Hostname'
[string]$WifiMacAddressColumn = 'Wifi MAC Adres'
[string]$ClassColumn = 'Klas'
[string]$CreatedUsersFile = 'AD-nieuwe-gebruikers.xlsx'

# Import and verify data
$InputData = $null
if ($InputFile -match '.csv')
{
    $InputData = Import-Csv -Path $InputFile -Delimiter ';'
}
else 
{
    $SheetName = (Get-ExcelSheetInfo -Path $InputFile)[0].Name
    $InputData = Import-Excel -Path $InputFile -WorksheetName $SheetName
}

# HostName and WifiMacAddress are required, abort if these aren't found
$ColumnNames = $InputData | Get-Member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name'
if ($ColumnNames -notcontains $HostNameColumn)
{
    Write-Error "Kolom `'$HostNameColumn`' zou de host name moeten bevatten, maar deze kolom bestaat niet."
    Exit
}

if ($ColumnNames -notcontains $WifiMacAddressColumn)
{
    Write-Error "Kolom `'$WifiMacAddressColumn`' zou het MAC-adres van de wifi-adapter moeten bevatten, maar deze kolom bestaat niet."
    Exit
}

[string]$WifiOU = 'OU=mac-wifi,OU=machines,OU=school,DC=campus,DC=romerocollege,DC=be'
[string]$Subgroup = 'mac-wifi'

Import-Module ActiveDirectory

Push-Location
Set-Location "AD:\$($WifiOU)" -ErrorAction Stop
$InputData | ? $WifiMacAddressColumn | % {
    [string]$AccountName = $_.$WifiMacAddressColumn.ToLower() -replace '-',''
    $ADuser = $null
    $ADuser = Get-ADUser $AccountName -Properties Description -ErrorAction SilentlyContinue
    if (-not $ADuser) {	
        $parms = @{
            Name = $AccountName;
            ChangePasswordAtLogon = $false;
            PasswordNeverExpires = $true;
            Enabled = $true;
            SamAccountName = $AccountName;
            AccountPassword = ConvertTo-SecureString -AsPlainText $AccountName -Force;
            UserPrincipalName = $AccountName;
            Description = $_.$HostNameColumn;
            DisplayName = $AccountName;
            CannotChangePassword = $true
        }
        Write-Host "Machine met MAC adres $($_.$WifiMacAddressColumn) toevoegen als $($_.$HostNameColumn)"
        New-ADUser @parms
        Add-ADGroupMember -Identity $Subgroup -Members $(Get-ADUser $AccountName)
    } else {
        if ($ADuser.Description -ne $_.$HostNameColumn) {
            Write-Host "Machine met MAC adres $($_.$WifiMacAddressColumn) aanpassen naar $($_.$HostNameColumn)"
            Set-ADUser -Description $_.$HostNameColumn -Identity $AccountName
        }
    }
}
Pop-Location