$PrinterName = "SECRWM-PR1"
if (((Get-DnsClientGlobalSetting).SuffixSearchList -contains "campus.romerocollege.be") -and -not (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue)) {
    Add-Printer -Name $PrinterName -DeviceURL "$PrinterName.campus.romerocollege.be"
}