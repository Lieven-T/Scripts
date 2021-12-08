Start-Transcript c:\temp\migrateto-ad.log

MkDir c:\temp
cscript //B slmgr.vbs /ipk NGWM3-3QCYC-YVF77-VJQV4-2KCJD
cscript //B slmgr.vbs /ato
$CurrentPath = $PSScriptRoot

Copy-Item "$CurrentPath\MigrateTo-AD-Phase2.ps1" c:\temp
$Action = New-ScheduledTaskAction -Execute 'Powershell.exe' -Argument "-ExecutionPolicy Bypass C:\Temp\MigrateTo-AD-Phase2.ps1"
$Trigger = New-ScheduledTaskTrigger -AtStartup
$Principal = New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
Register-ScheduledTask -TaskName "InstallADSoftware" -TaskPath "\" -Action $Action -Trigger $Trigger -Principal $Principal -ErrorAction Stop

[string]$userName = "campus\joiner"[string]$userPassword = "4jUXK278Gjr2"[string]$computerName = Read-Host -Prompt "Computernaam"[securestring]$secStringPassword = ConvertTo-SecureString $userPassword -AsPlainText -Force
[pscredential]$credObject = New-Object System.Management.Automation.PSCredential ($userName, $secStringPassword)
Add-Computer -DomainName campus.romerocollege.be -NewName $computerName -Credential $credObject -Restart $true

Stop-Transcript
