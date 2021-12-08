if ($env:COMPUTERNAME -match "\-\d+" -and (Get-Service "VeyonService" -ErrorAction SilentlyContinue)) {
    Set-Service -Name "VeyonService" -StartupType Automatic
}