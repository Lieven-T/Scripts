Param
(
    [Parameter()]
    $MacAddress,

    [Parameter(Position=1,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
    $HostName,

    [Parameter()]
    $Ip = '255.255.255.255'
)

function Wake {
    Param(
        [Parameter(Mandatory=$True,Position=1)][string]$MacAddress,
        [int]$port = 9
    )
    $Broadcast = [Net.IPAddress]::Parse($Ip)
 
    $MacOut = $MacAddress.ToUpper()
    $MacAddress = $MacAddress -replace '[-:\.]', ''
    $Target = 0,2,4,6,8,10 | % { [convert]::ToByte($MacAddress.substring($_,2),16) }
    $Packet = (,[byte]255 * 6) + ($Target * 16)
    
    $UDPclient = New-Object System.Net.Sockets.UdpClient
    $UDPclient.Connect($Broadcast,$Port)
    [void]$UDPclient.Send($Packet, 102) 
    Write-Output "Pakket verzonden naar $ip`:$port om '$MacOut' te wekken"
}

Write-Host $HostName
if (-not ($MacAddress -or $HostName)) {
    Write-Host "Gelieve één van beide parameters op te geven"
    Exit
}

if ($HostName -isnot [array]) { $HostName = @($HostName) }

if ($MacAddress) {
    if ($MacAddress -isnot [array]) { $MacAddress = @($MacAddress) }
    $MacAddress | % {
        Wake $_
    }
}

if ($HostName) {
    $HostName | % {
        Write-Host "Resolven van `'$_`'"
        $PostParams = @{ hostName = $_ }
        $WebResult = (Invoke-RestMethod -Uri http://orc-inventory/query -Method GET -Body $PostParams).computers
        if ($WebResult.Count) {
            $WebResult | % {
                Write-Host "Match: `'$($_.hostName)`' met MAC-adres `'$($_.macAddress)`'"
                Wake $_.macAddress
            }
        } else {
            Write-Host "Geen match voor $_"
        }
    }
}