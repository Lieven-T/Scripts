Start-Transcript c:\temp\migrateto-azuread.log
cscript //B slmgr.vbs /ipk NGWM3-3QCYC-YVF77-VJQV4-2KCJD
cscript //B slmgr.vbs /ato
Copy-Item \\orc-dc1\util\packages\desktop.ppkg c:\temp\desktop.ppkg
Copy-Item \\orc-dc1\util\scripts\MigrateTo-AzureAD-Phase2.ps1 c:\temp\MigrateTo-AzureAD-Phase2.ps1
Copy-Item \\orc-dc1\util\scripts\MigrateTo-AzureAD-Phase3.ps1 c:\temp\MigrateTo-AzureAD-Phase3.ps1
Remove-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultUserName -Force -ErrorAction SilentlyContinue
Remove-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultPassword -Force -ErrorAction SilentlyContinue
Remove-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name DefaultDomainName -Force -ErrorAction SilentlyContinue
Set-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name AutoAdminLogon -Value 0
Remove-LocalUser Kiosk -ErrorAction SilentlyContinue
Remove-LocalUser LocalAdmin -ErrorAction SilentlyContinue
Get-CimInstance win32_userprofile | ? { $_.LocalPath -notmatch 'admin' -and -not $_.Special} | Remove-CimInstance

Unregister-ScheduledTask -TaskName Afsluitenpc -Confirm:$false -ErrorAction SilentlyContinue
Unregister-ScheduledTask -TaskName Cleanup-Profiles -Confirm:$false -ErrorAction SilentlyContinue
Unregister-ScheduledTask -TaskName Create-LocalAdmin -Confirm:$false -ErrorAction SilentlyContinue
Unregister-ScheduledTask -TaskName "Disable Hibernate" -Confirm:$false -ErrorAction SilentlyContinue
Unregister-ScheduledTask -TaskName Get-ComputerInfo -Confirm:$false -ErrorAction SilentlyContinue
Unregister-ScheduledTask -TaskName Install-Laps -Confirm:$false -ErrorAction SilentlyContinue
Unregister-ScheduledTask -TaskName "Local admin account" -Confirm:$false -ErrorAction SilentlyContinue
Unregister-ScheduledTask -TaskName Set-AudioVideo -Confirm:$false -ErrorAction SilentlyContinue
Unregister-ScheduledTask -TaskName Shutdown-Upgrades -Confirm:$false -ErrorAction SilentlyContinue
Unregister-ScheduledTask -TaskName Update-Software -Confirm:$false -ErrorAction SilentlyContinue
Unregister-ScheduledTask -TaskName Update-Windows -Confirm:$false -ErrorAction SilentlyContinue

c:\ProgramData\chocolatey\choco.exe uninstall veyon

$Action = New-ScheduledTaskAction -Execute 'Powershell.exe' -Argument "-ExecutionPolicy Bypass C:\Temp\MigrateTo-AzureAD-Phase2.ps1"
$Trigger = New-ScheduledTaskTrigger -AtStartup
$Principal = New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
Register-ScheduledTask -TaskName "MigrateTo-AzureAD-Phase2" -Action $Action -Trigger $Trigger -Principal $Principal -ErrorAction Stop -TaskPath "\ORC"

[string]$userName = "campus\joiner"[string]$userPassword = "4jUXK278Gjr2"[securestring]$secStringPassword = ConvertTo-SecureString $userPassword -AsPlainText -Force
[pscredential]$credObject = New-Object System.Management.Automation.PSCredential ($userName, $secStringPassword)
Remove-Computer -Force -UnjoinDomainCredential $credObject -Restart
Stop-Transcript