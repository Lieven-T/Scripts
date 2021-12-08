$AccountName = "localadmin"
$Password = "Zomer2021!"
$HostName = $env:COMPUTERNAME
if (-not (Get-LocalUser -ErrorAction SilentlyContinue -Name $AccountName)) {
    $FileName = "c:\temp\createlocaladmin-$HostName.log"
    Start-Transcript $FileName

    Set-ExecutionPolicy Bypass -Force
    if (-not (Get-InstalledModule Posh-SSH -ErrorAction SilentlyContinue)) {
        Write-Host "Posh-SSH installeren..."
        Install-PackageProvider -Name "NuGet" -Force
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
        Install-Module Posh-SSH
    }
    Import-Module Posh-SSH

    Write-Host "Aanmaken lokale admin '$AccountName' op toestel $HostName"
    $HostName = $env:COMPUTERNAME
    New-LocalUser -Name $AccountName -Password (ConvertTo-SecureString $Password -Force -AsPlainText) -AccountNeverExpires -PasswordNeverExpires
    Add-LocalGroupMember -Group "Administrators" -Member $AccountName
    Write-Host "Account aangemaakt!"
    Stop-Transcript
    $Password = ConvertTo-SecureString "Boemtsjak123" -Force -AsPlainText
    $Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "mailer-laptops@romerocollege.be",$Password
    Send-MailMessage -From "mailer-laptops@romerocollege.be" -To "it-cvd@romerocollege.be","it-cnn@romerocollege.be" `
        -Subject "Instellen lokale admin $HostName" -Attachments $FileName -SmtpServer "smtp.office365.com" -Credential $Credentials -UseSsl

    $Password = ConvertTo-SecureString -AsPlainText "WqkvMcbPzl74BUIt_FamUJHidb4-u6ri" -Force
    $Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "intune",$Password
    $SFTPSession = New-SFTPSession -ComputerName '84.199.24.236' -Credential $credentials -AcceptKey:$true
    Set-SFTPItem -SessionId $SFTPSession.SessionId -Force -Path $FileName -Destination "/home/intune/logs"
    Remove-SFTPSession -SessionId $SFTPSession.SessionId
}