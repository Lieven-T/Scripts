if (-not (Test-Path 'C:\Program Files\LAPS')) {
	$SrcDir = "\\campus\util\software\_updates"
	$DestDir = "c:\temp\updates"

	$Arch = if ([Environment]::Is64BitProcess) { "x64" } else { "x86" }
	MkDir $DestDir
	Set-Location $DestDir

	Start-Transcript "c:\temp\install_laps.log"
    Write-Host "[LAPS] LAPS nog niet geïnstalleerd, wordt nu geïnstalleerd"
    Copy-Item "$SrcDir\LAPS.$Arch.msi" -Destination $DestDir
    msiexec /i "LAPS.$Arch.msi" /quiet

	Stop-Transcript
}

