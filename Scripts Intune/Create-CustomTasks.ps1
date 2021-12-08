$LogFile = "Create-CustomTasks_$($env:COMPUTERNAME).log"
Start-Transcript -Path "C:\temp\$LogFile"

if (-not (Get-ScheduledTask -TaskName Rename-Device -ErrorAction SilentlyContinue)) { 
    Write-Host "Get-ScheduledTask aanmaken..."

    if (-not (Get-InstalledModule Posh-SSH -ErrorAction SilentlyContinue)) {
        Write-Host "Posh-SSH installeren..."
        Install-PackageProvider -Name "NuGet" -Force
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
        Install-Module Posh-SSH
    }

    if (-not (Test-Path -Path c:\Windows\ORC -ErrorAction SilentlyContinue)) {
        MkDir c:\Windows\ORC
    }

    $Password = ConvertTo-SecureString -AsPlainText "WqkvMcbPzl74BUIt_FamUJHidb4-u6ri" -Force
    $Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "intune",$Password
    $SFTPSession = New-SFTPSession -ComputerName '84.199.24.236' -Credential $credentials -AcceptKey:$true
    Get-SFTPFile -SessionId $SFTPSession.SessionID -LocalPath C:\Windows\ORC -RemoteFile /home/intune/scripts/Shutdown-Upgrades.ps1 -Overwrite -ErrorAction Stop
    Get-SFTPFile -SessionId $SFTPSession.SessionID -LocalPath C:\Windows\ORC -RemoteFile /home/intune/scripts/Set-SleepSettings.ps1 -Overwrite -ErrorAction Stop

    $Action = New-ScheduledTaskAction -Execute 'Powershell.exe' -Argument "-ExecutionPolicy Bypass C:\Windows\ORC\Shutdown-Upgrades.ps1"
    $Trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Tuesday -At 10am
    $Principal = New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    Register-ScheduledTask -TaskName "Shutdown-Upgrades" -TaskPath "\ORC" -Action $Action -Trigger $Trigger -Principal $Principal -ErrorAction SilentlyContinue

    $Action = New-ScheduledTaskAction -Execute 'shutdown.exe' -Argument "-s -f -t 90"
    $Trigger = New-ScheduledTaskTrigger -Daily -At 8pm
    $Principal = New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    Register-ScheduledTask -TaskName "Shutdown-PC" -TaskPath "\ORC" -Action $Action -Trigger $Trigger -Principal $Principal -ErrorAction SilentlyContinue

    <#$Action = New-ScheduledTaskAction -Execute 'Powershell.exe' -Argument "-ExecutionPolicy Bypass C:\Windows\ORC\Set-SleepSettings.ps1"
    $Trigger = New-ScheduledTaskTrigger -AtStartup
    $Principal = New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    Register-ScheduledTask -TaskName "Set-SleepSettings" -TaskPath "\ORC" -Action $Action -Trigger $Trigger -Principal $Principal -ErrorAction SilentlyContinue#>

    Stop-Transcript
    Set-SFTPFile -SessionId $SFTPSession.SessionId -RemotePath "/home/intune/logs" -LocalFile "C:\temp\$LogFile" -Overwrite

    Remove-SFTPSession -SessionId $SFTPSession.SessionId
}