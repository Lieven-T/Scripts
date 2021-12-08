Start-Transcript c:\temp\Create-LocalAdmin.log

if ((Get-LocalUser administrator).Enabled) {
    Write-Host "Disable built in administrator"
	Disable-LocalUser Administrator
}

if (-not (Get-LocalUser -Name localadmin -ErrorAction Ignore)) {
    Write-Host "Create localadmin"
	$Password = ConvertTo-SecureString "%Jl<;CYN" -AsPlainText -Force
	New-LocalUser localadmin -AccountNeverExpires -FullName "Lokale Admin" -Password $Password
}

if (-not (Get-LocalGroupMember administrators | ? Name -Match "localadmin")) { 
    Write-Host "Local admin beheerder maken"
    Add-LocalGroupMember -Group Administrators -Member localadmin
}

if (-not (Get-LocalGroupMember administrators | ? Name -Match "localadmin")) { 
    Write-Host "Local admin beheerder maken"
    Add-LocalGroupMember -Group Administrators -Member localadmin
}

if (Get-LocalUser orcadmin -ErrorAction SilentlyContinue) { 
    Write-Host "Orcadmin scrappen" 
    Remove-LocalUser orcadmin
}

Stop-Transcript