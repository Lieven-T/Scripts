$Now = Get-Date
$Computers = (Get-ADComputer -Filter * -Property ms-Mcs-AdmPwd,ms-Mcs-AdmPwdExpirationTime | ? { 
    $Date = $([datetime]::FromFileTime([convert]::ToInt64($_."ms-Mcs-AdmPwdExpirationTime")))
    ((-not $_.'ms-Mcs-ADmPwd') -or ($Date -lt $Now)) -and $_.DistinguishedName -notmatch "-LT|ORC-|H004-|LAPTOP-|-DC"
})
$Computers | ft
Write-Host "Totaal: $($Computers.Count)"