if ((Get-LocalUser administrator).Enabled) {
	Disable-LocalUser Administrator
}

if (-not (Get-LocalUser -Name localadmin -ErrorAction Ignore)) {
	$Password = ConvertTo-SecureString "%Jl<;CYN" -AsPlainText -Force
	New-LocalUser localadmin -AccountNeverExpires -FullName "Lokale Admin" -Password $Password
}

if (-not (Get-LocalGroupMember administrators | ? Name -Match "localadmin")) { Add-LocalGroupMember -Group Administrators -Member localadmin }

if (-not (Get-LocalGroupMember administrators | ? Name -Match "localadmin")) { Add-LocalGroupMember -Group Administrators -Member localadmin }

if (Get-LocalUser orcadmin -ErrorAction SilentlyContinue) { Remove-LocalUser orcadmin }
