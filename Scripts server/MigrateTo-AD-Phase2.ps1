Start-Transcript c:\temp\migrateto-ad-phase2.log

GpUpdate /force
Start-ScheduledTask Get-ComputerInfo
Start-ScheduledTask Update-Software
Start-ScheduledTask Update-Windows
Unregister-ScheduledTask -TaskName "InstallADSoftware" -Confirm:$false

Stop-Transcript