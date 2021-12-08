$LogFiles = Get-Item -Path \\campus\util\bios\log\*
$ReportFiles = Get-Item -Path \\campus\util\bios\report\*

$MissingReports = $Logfiles | ? { ($_.Name -split "_")[0] + ".txt" -notin ($ReportFiles | select -ExpandProperty name ) } | % { ($_.Name -split '_')[0] }
$FaultySettings = $ReportFiles | ? { (Get-Content $_ | Out-String) -match "Tuesday\s*\*Disable" } | % { ($_.Name -split "\." )[0] }

Write-Host "`n=====================`nOntbrekende rapporten`n====================="
$MissingReports
Write-Host "`n`n`n==================`nFoute instellingen`n=================="
$FaultySettings
