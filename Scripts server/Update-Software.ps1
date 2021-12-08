Start-Transcript "\\campus\util\transcripts\$($env:COMPUTERNAME)-updates.log"

function Compare-Versions {
    param (
        [Parameter(Position=1)][String] $OldVersion,
        [Parameter(Mandatory=$True,Position=2)][String] $NewVersion
    )

    if (-not $OldVersion) { return $true }

    $OldVersionArray = $OldVersion.Split('-.,')
    $NewVersionArray = $NewVersion.Split('-.,')

    $Size = if ($OldVersionArray.Length -gt $NewVersionArray.Length) { $OldVersionArray.Length } Else { $NewVersionArray.Length }
    $Sizing = 10000
    $Decimals = [Math]::Pow($Sizing, $Size -1)
    $NewVersionNumber = 0
    $OldVersionNumber = 0
    For ($i = 0;$i -lt $Size;$i++) {
        if ($i -lt $NewVersionArray.Length) { $NewVersionNumber += [int]::Parse($NewVersionArray[$i]) * $Decimals }
        if ($i -lt $OldVersionArray.Length) { $OldVersionNumber += [int]::Parse($OldVersionArray[$i]) * $Decimals }
        $Decimals /= $Sizing
    }

    return $NewVersionNumber -gt $OldVersionNumber
}

function Get-MSIVersion {
    param (
        [IO.FileInfo] $MSI
    )
 
    if (!(Test-Path $MSI.FullName)) {
        throw "File '{0}' does not exist" -f $MSI.FullName
    }
 
    try {
        $windowsInstaller = New-Object -com WindowsInstaller.Installer
        $database = $windowsInstaller.GetType().InvokeMember(
            "OpenDatabase", "InvokeMethod", $Null,
            $windowsInstaller, @($MSI.FullName, 0)
        )
 
        $q = "SELECT Value FROM Property WHERE Property = 'ProductVersion' AND Value IS NOT NULL"
        $View = $database.GetType().InvokeMember(
            "OpenView", "InvokeMethod", $Null, $database, ($q)
        )
 
        $View.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $View, $Null)
        $record = $View.GetType().InvokeMember( "Fetch", "InvokeMethod", $Null, $View, $Null )
        $version = $record.GetType().InvokeMember( "StringData", "GetProperty", $Null, $record, 1 )

        return [string]$version
    } catch {
        throw "Failed to get MSI file version: {0}." -f $_
    }
}

function Update-Software {
    param(
        [Parameter(Mandatory=$True)][String] $Program, 
        [Parameter(Mandatory=$True)][ValidateSet("msi","versioned","named","script")] [String] $SourceType, 
        [String]$Params,
        [ValidateSet("package","manual")][String]$DestType = "package",
        [String]$ExePath,
        [bool]$Force = $true,
        [String]$Target,
        [bool]$NoVersionCheck = $false,
        [bool]$Reinstall = $false
    )

    Set-Location $SrcDir
    $Program = $Program.ToUpper()
    
    if ($env:COMPUTERNAME -notmatch $Target -and $env:COMPUTERNAME -ne "PC-TROLI2") { 
        Write-Host "[$Program] Computer is geen doel" 
        Return
    }

    $OldVersion = If ($DestType -eq "package") { 
        $ProgramList = ($InstalledPrograms | ? Name -like "$Program*" | Sort-Object -Property Version -Descending)
        If ($ProgramList) {
            if ($NoVersionCheck) { "1000" } else { $ProgramList[0].Version }
        } Else {
            $null
        }
    } Else {
        $File = 
            If (Test-Path $ExePath -ErrorAction SilentlyContinue) {
                Get-Item $ExePath
            } ElseIf (Test-Path ($OldFile = "C:\Program Files\$program\$program.exe") -ErrorAction SilentlyContinue) { 
                Get-Item $OldFile
            } ElseIf (Test-Path($OldFile = "C:\Program Files (x86)\$program\$program.exe") -ErrorAction SilentlyContinue) {
                Get-Item $OldFile
            } Else {
                $null
            }
        if ($File) {
            if ($NoVersionCheck) { "1000" } else { $File.VersionInfo.ProductVersion }
        }  else {
            $null
        }
    }

    $NewFile = $null
    $NewVersion = $null
    $NewFileItem = $null
    Switch($SourceType) {
         "msi" { 
            $NewFile = "$SrcDir\$Program.msi"
            [String]$NewVersion = Get-MSIVersion($NewFile)
        }
         "versioned" { 
            $NewFile = "$SrcDir\$Program.exe"
            [String]$NewVersion = (Get-Item $NewFile).VersionInfo.Productversion
        }
         "named" { 
            $NewFileItem = Get-ChildItem -Path $SrcDir -Filter "$Program.*"
            $NewFile = $NewFileItem.FullName
            [String]$NewVersion = $NewFileItem.Name.Replace(".exe","").ToUpper().Replace("$Program.","") 
        }
        "script" {
            [String]$NewVersion = & "$SrcDir\$Program\$Program.ps1" -Action version
        }
    }
    
    if (($OldVersion -eq $null) -and -not $Force) { 
        Write-Host "[$Program] Programma niet geïnstalleerd" 
        Return
    }

    $NewVersion = $NewVersion.Trim()
    Write-Host "[$Program] Versies: $OldVersion (oud); $NewVersion (nieuw)"
    if (Compare-Versions $OldVersion $NewVersion) {
        Write-Host "[$Program] Update wordt geïnstalleerd"
        If ($Reinstall -and $DestType -eq "package") {
            Write-Host "[$Program] Oude versie verwijderen"
            $InstalledPrograms | ? Name -like "$Program*" | Uninstall-Package
        }
        If ($SourceType -eq "script") {
            & "$SrcDir\$Program\$Program.ps1"
        } else {
            Copy-Item $NewFile -Destination $DestDir
            Switch($SourceType) {
                "msi" {  & msiexec /i "$Program.msi" $Params.Split() /qn /norestart | Out-Host }
                "versioned" { & ".\$Program.exe" $params.Split() | Out-Host }
                "named" { & ".\$($NewFileItem.Name)" $params.Split() | Out-Host }
            }
        }
    } else {
        Write-Host "[$Program] Update niet nodig"
    }
}

################
# BASIS CHECKS #
################

If ((Get-CimInstance -ClassName Win32_OperatingSystem).ProductType -ne 1) {
    Write-Host "Deze machine is een server: uitvoering stoppen"
    Exit
}

If (($env:COMPUTERNAME -match "-SG") -or ($env:COMPUTERNAME -match "CAMERA")) {
    Write-Host "Deze machine is een appliance: uitvoering stoppen"
    Exit
}

Push-Location

$SrcDir = "\\campus\util\software\_updates"
$DestDir = "c:\temp\updates"

RmDir $DestDir -Force -Recurse -ErrorAction SilentlyContinue
MkDir $DestDir
Set-Location $DestDir

$InstalledPrograms = Get-Package


##########
# OFFICE #
##########

if (-not ($InstalledPrograms | ? Name -match "Office 16")) {
    Write-Host "OFFICE INSTALLEREN..."
    Push-Location
    Copy-Item $SrcDir\Office $DestDir -Recurse
    Set-Location $DestDir\Office
    .\setup.exe /download CVD.xml
    .\setup /configure CVD.xml
    Pop-Location
} else {
    Write-Host "Office geïnstalleerd"
}


######################
# INSTALL CHOCOLATEY #
######################

if(-not(Test-Path C:\ProgramData\chocolatey\choco.exe)){
    Write-Host "Seems Chocolatey is not installed, installing now"
    Remove-Item C:\ProgramData\chocolatey -Recurse -Force
    & iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
    C:\ProgramData\chocolatey\choco.exe feature enable -n allowGlobalConfirmation
} else {
    Write-Host "Chocolatey installed!"
}


###########
# UPDATES #
###########

Get-Package | ? Name -match "Acrobat Reader" | ? Name -NotMatch "MUI" | Uninstall-Package -Force -Confirm:$false
Update-Software -Program "Avigilon" -SourceType msi -Target "SECRWM\-LK|SECRWM2\-LK|JEFF\-LT"
Update-Software -Program "eMindMaps" -SourceType named -Params "/S"
Update-Software -Program "Winfakt" -SourceType script -DestType manual -ExePath "C:\WinFakt-Scholen\WinFakt.exe" -NoVersionCheck $true -Target "H303"
Update-Software -Program "BOB 50 SQL" -SourceType script -Target "H302"
Update-Software -Program "Chemsketch" -SourceType script -DestType manual -ExePath "C:\ACDFREE12\CHEMSK.EXE" -Target "F116-LK"
Update-Software -Program "graphmatica" -SourceType msi -Target "-LK"
Update-Software -Program "sprint" -SourceType script -DestType manual -ExePath "C:\Program Files (x86)\SprintPlus 3\sprint.exe" -Target "G002|SECREIT\-LT|STUDIOCVD|JEFF\-LT"
Update-Software -Program "TI-SmartView CE-T" -SourceType script -Target "-LK"

c:\ProgramData\Chocolatey\choco.exe upgrade googlechrome firefox adobereader microsoft-teams.install libreoffice 7zip gimp audacity libreoffice notepadplusplus openshot vlc paint.net r r.studio laps geogebra-classic.install --noprogress
c:\ProgramData\Chocolatey\choco.exe pin add -n=googlechrome
c:\ProgramData\Chocolatey\choco.exe pin add -n=firefox
c:\ProgramData\Chocolatey\choco.exe pin add -n=adobereader
c:\ProgramData\Chocolatey\choco.exe pin add -n=microsoft-teams.install
if ($env:COMPUTERNAME -match "F114|F215|F303|G002|G201|G204|H002|H003|H101|H107|H203|H204|H301|H302|H303|H306|H307|SECREIT\-LT|SECRWM\-LT") {
    if ($env:COMPUTERNAME -match "LK") {
        c:\ProgramData\Chocolatey\choco.exe upgrade veyon --params '"/Master /Config:\\campus\util\software\_updates\Veyon\veyon.json"' --noprogress
    } else {
        c:\ProgramData\Chocolatey\choco.exe upgrade veyon --params '"/Config:\\campus\util\software\_updates\Veyon\veyon.json"' --noprogress
    }
}

# Update-Software -Program "Belgium e-ID middleware" -SourceType msi -Target "-LK" -NoVersionCheck $true

Pop-Location

Stop-Transcript