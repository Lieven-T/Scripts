Start-Transcript c:\temp\migrateto-azuread-phase2.log
Unregister-ScheduledTask -TaskName "MigrateTo-AzureAD-Phase2" -Confirm:$false

<#$Action = New-ScheduledTaskAction -Execute 'Powershell.exe' -Argument "-ExecutionPolicy Bypass C:\Temp\MigrateTo-AzureAD-Phase3.ps1"
$Trigger = New-ScheduledTaskTrigger -AtStartup
$Principal = New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
Register-ScheduledTask -TaskName "MigrateTo-AzureAD-Phase3" -TaskPath "\ORC" -Action $Action -Trigger $Trigger -Principal $Principal#>

Install-ProvisioningPackage -PackagePath c:\temp\desktop.ppkg -ForceInstall -QuietInstall
Stop-Transcript