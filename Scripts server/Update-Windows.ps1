# ============================== ##
# Update/Install PSWindowsUpdate ##
# ============================== ##

if (-not (Get-Module PSWindowsUpdate)) {
    Install-PackageProvider -Name Nuget -Force
    Install-Module -Name PSWindowsUpdate  -Force
} else {
    Update-Module -Name PSWindowsUpdate -Force
}


# =============== #
# Regular updates #
# =============== #

Import-Module PSWindowsUpdate
Get-WUInstall -ForceInstall


# =============== #
# Feature updates #
# =============== #

if ((gcim Win32_OperatingSystem).Version -lt '10.0.19044') {
    Start-Transcript c:\temp\Update-Windows.log
    Write-Output "Windows Upgrade uitvoeren"
    $Dir = "c:\temp"
    Copy-Item "\\campus\util\software\Windows10Upgrade.exe" "C:\temp\windows10upgrade.exe" -ErrorAction SilentlyContinue -Force
    Start-Process -FilePath "C:\temp\windows10upgrade.exe" -ArgumentList "/quietinstall /skipeula /auto upgrade /copylogs $Dir"
    Stop-Transcript
}