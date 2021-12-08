Start-Transcript c:\temp\Cleanup-Profiles.log

$UserProfiles = Get-CimInstance win32_userprofile
Write-Host "=== GEVONDEN PROFIELEN ==="
$UserProfiles | select LocalPath,LastUseTime | Sort LocalPath | Write-Host

$OldUserProfiles = $UserProfiles | Where {$_.LastUseTime -lt $(Get-Date).Date.AddDays(-60) -and $_.LocalPath -notmatch 'admin' -and -not $_.Special}
Write-Host "`n=== OP TE KUISEN PROFIELEN ==="
$OldUserProfiles | select LocalPath,LastUseTime | Sort LocalPath | Write-Host
$OldUserProfiles  | Remove-CimInstance

Stop-Transcript