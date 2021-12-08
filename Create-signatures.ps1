Param
(
    [Parameter(Mandatory=$True,Position=1)]
    [string]$InputFile
)

[string]$FirstNameColumn = 'Voornaam'
[string]$LastNameColumn = 'Achternaam'
[string]$FaxColumn = 'Faxnr'
[string]$TelColumn = 'Telnr'
[string]$MobColumn = 'Gsmnr'
[string]$JobColumn = 'Functie'
[string]$DomainColumn = 'Domeinschool'
[string]$AddressColumn = 'Adres'
[string]$DomainName = 'romerocollege.be'

$DomainLogos = @{
    'WM' = 'http://www.romerocollege.be/_extra_/email_orc_wm.png';
    'TEW' = 'http://www.romerocollege.be/_extra_/email_orc_tew.png'; 
    'EIT' = 'http://www.romerocollege.be/_extra_/email_orc_eit.png'; 
    'TE' = 'http://www.romerocollege.be/_extra_/email_orc_te.png'; 
    '1G' = 'http://www.romerocollege.be/_extra_/email_orc_1egr.png'; 
    'BUBO' = 'http://www.romerocollege.be/_extra_/email_orc_blo.png'; 
    'BO' = 'http://www.romerocollege.be/_extra_/email_orc_basis.png';
    'CLW' = 'http://www.romerocollege.be/_extra_/email_orc_lw.png'; 
    'ORS' = 'http://www.romerocollege.be/_extra_/_email_ors.png'; 
    'Default' = 'http://www.romerocollege.be/_extra_/email_orc_sec.png'
}

$Template = 
'<table style="font-size:8.0pt; font-family:Helvetica,sans-serif; color:#575756;line-height:1.1">
    <tr>
        <td style="width: 188px; border:none;border-right:solid #575756 1.0pt;padding:0; padding-right:10px; vertical-align: top; text-align: left" rowspan="3">
            <a href="http://www.romerocollege.be/" style="display:block">
                <img width="184" style="width: 184px;margin-top:7px" alt="&Oacute;scar Romeroscholen vzw" src="{LogoSrc}">
            </a>
        </td>
        <td height="22" style="font-size:10.0pt;padding-left:10px;"><div style="margin-top:5px">{Name} - {JobTitle}</div></td>
    </tr>
    <tr>
        <td style="vertical-align: bottom;padding-left:10px;line-height:1.3">
            <div>&Oacute;scar Romeroscholen VZW</div>
            <div>{Address}</div>
            <table style="font-size:8.0pt; font-family:Helvetica,sans-serif; color:#575756;line-height:1.3" border="0" cellpadding="0" cellspacing="0">
                {Tel}
                {Mob}
                {Fax}
            </table>
            <a href="mailto:{Email}" style="color:blue">{Email}</a>&nbsp;|&nbsp;<a href="http://www.romerocollege.be/" style="color:#575756">www.romerocollege.be</a>
        </td>
    </tr>
    <tr>
        <td style="vertical-align: bottom;padding-left:10px;line-height:1.3">
            <br>
            <div>Ondernemingsnummer: 0415819204</div>
            <div>Ondernemingsrechtbank Gent afdeling Dendermonde</div>
        </td>
    </tr>
    <tr><td height="30" style="height:30px; vertical-align:bottom;"><a style="text-decoration: none; color: #aaa" href="http://www.romerocollege.be/algemeen/disclaimer">disclaimer</a></td><td></td></tr>
</table>'

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

    # Compose email address as firstname.lastname@domainname. Diacritics are filtered out and replaced by plain ASCII characters
	$Email = "$(Remove-Diacritics $_.$FirstNameColumn).$(Remove-Diacritics $_.$LastNameColumn)@$($DomainName)"

	# Strip all spaces, single quotes and dashes
    $Email = $Email -replace '[-\s'''’]',''

    $Account | Add-Member -Name "Email" -Value $Email -MemberType NoteProperty
}

# Import and verify data
$InputData = $null
if ($InputFile -match '.csv')
{
    $InputData = Import-Csv -Path $InputFile
}
else 
{
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

if ($ColumnNames -notcontains $JobColumn)
{
    Write-Error "Kolom `'$JobColumn`' zou de functieomschrijveing moeten bevatten, maar deze kolom bestaat niet."
    Exit
}

# When there are user to create, run through the list and process each item
$InputData | ForEach-Object {

    Create-Email $_
    $FileName = Remove-Diacritics "$($_.$FirstNameColumn)-$($_.$LastNameColumn).html"
    Write-Host "Aanmaken signaturebestand $FileName"
        
	$Output = $Template
    $Output = $Output -replace '{Name}', "$($_.$FirstNameColumn.ToUpper()) $($_.$LastNameColumn.ToUpper())"
    $Output = $Output -replace '{JobTitle}', ($_.$JobColumn -replace '&', '&amp;')
    $Output = $Output -replace '{Address}', $_.$AddressColumn
    $Output = $Output -replace '{Email}', $_.Email.ToLower()
        
    if ($_.$FaxColumn.Length -gt 0) {
        $Output = $Output -replace '{Fax}', "<tr><td width=`"15`">F</td><td>$($_.$FaxColumn)</td></tr>"
    } 
    else 
    {
        $Output = $Output -replace '{Fax}', ''
    }
    
    if ($_.$MobColumn.Length -gt 0) 
    {
        $Output = $Output -replace '{Mob}', "<tr><td width=`"15`">M</td><td>$($_.$MobColumn)</td></tr>"
    }
    else
    {
        $Output = $Output -replace '{Mob}', ''
    }

    if ($_.$TelColumn.Length -gt 0) 
    {
        $Output = $Output -replace '{Tel}', "<tr><td width=`"15`">T</td><td>$($_.$TelColumn)</td></tr>"
    }
    else
    {
        $Output = $Output -replace '{Tel}', ''
    }

    if ($DomainLogos.ContainsKey($_.$DomainColumn))
    {
        $Output = $Output -replace '{LogoSrc}', $DomainLogos[$_.$DomainColumn]
        $Output = $Output -replace '{Spacer}', '<div style="font-size:12px">&nbsp;</div>'
    }
    else
    {
        $Output = $Output -replace '{LogoSrc}', $DomainLogos['Default']
        $Output = $Output -replace '{Spacer}', ''
    }

	$Output | Out-File -FilePath $FileName -Encoding utf8
}
