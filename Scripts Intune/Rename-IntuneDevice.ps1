Set-ExecutionPolicy Bypass -Force
if (-not (Get-InstalledModule ImportExcel -ErrorAction SilentlyContinue)) {
    Write-Host "ImportExcel installeren..."
    Install-PackageProvider -Name "NuGet" -Force
    Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
    Install-Module ImportExcel
}

if (-not (Get-InstalledModule Posh-SSH -ErrorAction SilentlyContinue)) {
    Write-Host "Posh-SSH installeren..."
    Install-PackageProvider -Name "NuGet" -Force
    Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
    Install-Module Posh-SSH
}
Import-Module Posh-SSH

#Get serial number of Laptop
$SerNum = (Get-WmiObject win32_bios).SerialNumber.ToString()
$OldName = $env:computername

$Password = ConvertTo-SecureString -AsPlainText "WqkvMcbPzl74BUIt_FamUJHidb4-u6ri" -Force
$Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "intune",$Password

$SFTPSession = New-SFTPSession -ComputerName '84.199.24.236' -Credential $credentials -AcceptKey:$true

Get-SFTPItem -SessionId $SFTPSession.SessionID -Destination c:\temp -Path /home/intune/scripts/devicelist.xlsx -Force
$data = Import-Excel C:\temp\devicelist.xlsx

$Computer = $data | ? Serienummer -eq $SerNum
if ($Computer) {
    if ($OldName -ne $Computer.Label) {
        $LogFile = "Rename-IntuneDevice_$($env:COMPUTERNAME).log"
        Start-Transcript -Path "C:\temp\$LogFile"

        Write-Host "Computernaam wijzigen..."
        Write-Host "    Serienummer: $SerNum"
        Write-Host "    Oude naam: $OldName"
        $NewName = $Computer.Label
        Write-Host "    Nieuwe naam: $NewName"
        Rename-Computer -NewName $NewName

        Stop-Transcript
        Set-SFTPItem -SessionId $SFTPSession.SessionId -Force -Path "C:\temp\$LogFile" -Destination "/home/intune/logs"
    }
}

Remove-SFTPSession -SessionId $SFTPSession.SessionId