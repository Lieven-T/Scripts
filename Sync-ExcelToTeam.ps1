$ClassTeam = New-Team -DisplayName "2021_2OK" -MailNickName "2021_2OK" -Template EDU_Class -AllowCreatePrivateChannels $false

$inputdata = Import-Excel .\2OKteam.xlsx
$InputData | ? Rol -EQ "lid" | % { Add-TeamUser -GroupId $ClassTeam.GroupId -User "$($_.Naam)@student.romerocollege.be" -Role Member }
$InputData | ? Rol -EQ "eigenaar" | % { Add-TeamUser -GroupId $ClassTeam.GroupId -User "$($_.Naam)@romerocollege.be" -Role Owner }