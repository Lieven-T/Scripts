Start-Transcript c:\temp\migrateto-azuread-phase3.log
Unregister-ScheduledTask -TaskName "MigrateTo-AzureAD-Phase3" -Confirm:$false

if (Test-Connection orc-inventory -Count 1 -ErrorAction SilentlyContinue) {
    Write-Host 'Inventory server gevonden'
    $SerNum = (Get-WmiObject win32_bios).SerialNumber.ToString()
    $Data = Invoke-RESTMethod -Uri "http://orc-inventory/query?serialNumber=$SerNum" -UseBasicParsing
    if ($Data.computers) {
        Write-Host "Is Inventory-device"
        $NewName = $Data.computers[0].hostname
        if ($OldName -ne $NewName) {
            Write-Host "Nieuwe naam: $NewName"
            Rename-Computer -NewName $NewName -Restart:$true
        }
    }
}
Stop-Transcript