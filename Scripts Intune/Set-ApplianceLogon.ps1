Start-Transcript c:\temp\appliance.log
New-LocalUser -Password (ConvertTo-SecureString -Force -AsPlainText -String 'appliance') -AccountNeverExpires -Name appliance -UserMayNotChangePassword -PasswordNeverExpires -ErrorAction SilentlyContinue
Set-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultUserName -Value 'appliance' -ErrorAction SilentlyContinue
Set-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultPassword -Value 'appliance' -ErrorAction SilentlyContinue
Set-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultDomainName -Value '.' -ErrorAction SilentlyContinue
Set-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name AutoAdminLogon -Value 1 -ErrorAction SilentlyContinue

if (-not (Test-Path -Path c:\Windows\ORC -ErrorAction SilentlyContinue)) {
    MkDir c:\Windows\ORC
}

$Password = ConvertTo-SecureString -AsPlainText "WqkvMcbPzl74BUIt_FamUJHidb4-u6ri" -Force
$Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "intune",$Password
$SFTPSession = New-SFTPSession -ComputerName '84.199.24.236' -Credential $credentials -AcceptKey:$true
Get-SFTPFile -SessionId $SFTPSession.SessionID -LocalPath C:\Windows\ORC -RemoteFile /home/intune/scripts/kiosk-eit.lnk -Overwrite -ErrorAction Stop
Get-SFTPFile -SessionId $SFTPSession.SessionID -LocalPath C:\Windows\ORC -RemoteFile /home/intune/scripts/kiosk-wm.lnk -Overwrite -ErrorAction Stop

Remove-SFTPSession -SessionId $SFTPSession.SessionId
Stop-Transcript