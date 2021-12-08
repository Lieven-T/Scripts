$TempFolder = "c:\temp"
$DfsFolder = "\\campus\util"
New-Item -ErrorAction Ignore -ItemType directory -Path $TempFolder

### ================== ###
### GET AUTOPILOT INFO ###
### ================== ###

$AutoPilotHash = $null
try {
    Set-Location -Path "C:\Temp"
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted -Force:$true
    if (-not (Get-Command Get-WindowsAutoPilotInfo -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name "NuGet" -Force
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
        Install-Script -Name Get-WindowsAutoPilotInfo -Force
    }

    & 'C:\Program Files\WindowsPowerShell\Scripts\Get-WindowsAutoPilotInfo.ps1' -OutputFile "autopilot_hash.csv"
    $AutoPilotHash = (Import-Csv -Path C:\temp\autopilot_hash.csv)[0].'Hardware Hash'
} catch {
}


### ======================== ###
### REGISTER COMPUTER CONFIG ###
### ======================== ###

$Processor = (Get-CimInstance -ClassName Win32_Processor).Name

$ComputerInfo = Get-CimInstance -ClassName Win32_ComputerSystem
$RAM = [math]::Round($ComputerInfo.TotalPhysicalMemory / 1024 / 1024 / 1024)
$Model = $ComputerInfo.Model

$SerialNumber = (Get-CimInstance -ClassName Win32_BIOS).SerialNumber

$Disk = (Get-Disk)[0] | ? IsBoot
$DiskSize = [math]::Round($Disk.Size / 1024 / 1024 / 1024)
$DiskSerialNumber = $Disk.SerialNumber
if ($DiskSerialNumber) { $DiskSerialNumber = $DiskSerialNumber.Trim() }
$MediaType = (Get-PhysicalDisk | ? SerialNumber -eq $DiskSerialNumber).MediaType

$MacAddress = $null
$EthernetAdapter = Get-NetAdapter -Physical -Name '*ethernet*'
if ($EthernetAdapter) { $MacAddress = $EthernetAdapter[0].MacAddress }

$WifiMacAddress = $null
$WifiAdapter = Get-NetAdapter -Physical -Name '*wi*fi*'
if ($WifiAdapter) { $WifiMacAddress = $WifiAdapter[0].MacAddress }

$HostName = $env:computername
$OSInfo = gcim Win32_OperatingSystem | select Version, InstallDate

$WmiDisk = Get-WMIObject Win32_LogicalDisk | ? DeviceId -EQ "C:"
$FreeSpace = [math]::Round($WmiDisk.FreeSpace / 1024 / 1024 / 1024)

$Errors = @()
if (Get-Service | ? DisplayName -match "G Data") { $Errors += "GData" }
if ($FreeSpace -lt 0.1) { $Errors += "Schijfruimte" }
if (-not (Test-Path 'C:\Program Files (x86)\Teams Installer')) { $Errors += "Teams" }
$ErrorString = $Errors | % { "$_" }

$PostParams = @{
                    hostName = $HostName;
                    macAddress = $MacAddress;
                    wifiMacAddress = $WifiMacAddress;
                    mediaType = $MediaType;
                    processor = $Processor;
                    diskSize = $DiskSize;
                    model=$Model;
                    ramSize = $RAM;
                    serialNumber = $SerialNumber;
                    osVersion = $OSInfo.Version;
                    installDate = $OSInfo.InstallDate.ToString("yyyyMMddHHmmss");
                    freeSpace = $FreeSpace
                    errors = $ErrorString
                    autoPilotHash = $AutoPilotHash
                }
Invoke-WebRequest -Uri http://orc-inventory/register -Method POST -Body $PostParams -UseBasicParsing


### ================= ###
### GET BIOS SETTINGS ###
### ================= ###

$PasswordFile = "password.bin"

$ReportFile = "$HostName.txt"
$LogFile = "$($HostName)_log.txt"

if ($Model -match "7900") {

    $SettingsFile = "settings_old.txt"
    New-Item -ErrorAction Ignore -ItemType directory -Path $TempFolder
    Copy-Item "$DfsFolder\scripts\BiosConfigUtility64_old.exe" $TempFolder
    Set-Location $TempFolder

    if (-not (Test-Path -Path "$DfsFolder\bios\log\$LogFile")) {
        Copy-Item "$DfsFolder\bios\$SettingsFile" $TempFolder -Force -ErrorAction Ignore
        Remove-Item $TempFolder\$($HostName)_log*.txt -Force

        .\BiosConfigUtility64_old.exe /setconfig:$SettingsFile > $LogFile
        Copy-Item -Path "$LogFile" -Destination "$DfsFolder\bios\log" -Force
    }


    .\BiosConfigUtility64_old.exe /getconfig:$ReportFile
    Copy-Item $ReportFile "$DfsFolder\bios\report\" -Force

} else {

    $SettingsFile = "settings.txt"
    New-Item -ErrorAction Ignore -ItemType directory -Path $TempFolder
    Copy-Item "$DfsFolder\scripts\BiosConfigUtility64.exe" $TempFolder
    Set-Location $TempFolder

    if (-not (Test-Path -Path "$DfsFolder\bios\log\$LogFile")) {
        Copy-Item "$DfsFolder\bios\$SettingsFile" $TempFolder -Force -ErrorAction Ignore
        Copy-Item "$DfsFolder\bios\$PasswordFile" $TempFolder -Force -ErrorAction Ignore
        Remove-Item $TempFolder\$($HostName)_log*.txt -Force

        .\BiosConfigUtility64.exe /set:$SettingsFile /npwdfile:"`"`"" /logpath:"$TempFolder\$LogFile" > $LogFile
        Copy-Item -Path "$LogFile" -Destination "$DfsFolder\bios\log" -Force
    }


    .\BiosConfigUtility64.exe /get:$ReportFile
    Copy-Item $ReportFile "$DfsFolder\bios\report\" -Force

}
