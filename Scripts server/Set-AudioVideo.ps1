if ($env:computername -match '^(F|G|H)\d{3}-LK\d?$')
{
	$AudioDevices = Get-PnpDevice | Where-Object { $_.PNPClass -eq 'MEDIA' -and $_.FriendlyName -match 'High Definition Audio' }
	if ($AudioDevices | Where-Object { $_.FriendlyName -match 'Realtek' })
	{
		$AudioDevices | Where-Object { $_.FriendlyName -notmatch 'Realtek' -and $_.Status -eq 'OK' } | Disable-PnpDevice -Confirm:$false
	}
	else
	{
		$AudioDevices | Where-Object { $_.FriendlyName -notmatch 'Realtek' -and $_.Status -eq 'Error' } | Enable-PnpDevice -Confirm:$false
	}
}