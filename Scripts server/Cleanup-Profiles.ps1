$UserProfiles = Get-CimInstance win32_userprofile
$UserProfiles | select LocalPath,LastUseTime | Sort LocalPath | Write-Host

$OldUserProfiles = $UserProfiles | Where {$_.LastUseTime -lt $(Get-Date).Date.AddDays(-60) -and $_.LocalPath -notmatch 'admin' -and -not $_.Special}
$OldUserProfiles | select LocalPath,LastUseTime | Sort LocalPath | Write-Host
$OldUserProfiles  | Remove-CimInstance
