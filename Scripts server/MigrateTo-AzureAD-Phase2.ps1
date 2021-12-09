Start-Transcript c:\temp\migrateto-azuread-phase2.log
Unregister-ScheduledTask -TaskName "MigrateTo-AzureAD-Phase2" -Confirm:$false

Install-ProvisioningPackage -PackagePath c:\temp\desktop.ppkg -ForceInstall -QuietInstall
Stop-Transcript