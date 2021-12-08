# Andere groepen: romerocollege_ en niet "leerkrachten" of "leerlingen" -> lijst tonen van groepen; licenties schrappen

$licensePlanList = Get-AzureADSubscribedSku
#$LicenseToAdd = $licensePlanList | ? SkuPartNumber -eq "M365EDU_A3_FACULTY"
$LicenseToRemove = $licensePlanList | ? SkuPartNumber -eq "M365EDU_A3_FACULTY"
$Licenses = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses

$AssignedLicenseToAdd = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
$AssignedLicenseToAdd.SkuId = $LicenseToAdd.SkuId
$Licenses.AddLicenses = $AssignedLicenseToAdd

$Licenses.RemoveLicenses = $LicenseToRemove.SkuId

$users = Import-Excel .\users.xlsx
$speciallekes | % {
    $userUPN = $_.UserPrincipalName
    Write-Host $userUPN
    Set-AzureADUserLicense -ObjectId $userUPN -AssignedLicenses $licenses
}

# Caroline De Backer
# Marleen Joos

Get-AzureAdUser -All $true | % {
    $licensed=$False ; For ($i=0; $i -le ($_.AssignedLicenses | Measure).Count ; $i++) { If( [string]::IsNullOrEmpty(  $_.AssignedLicenses[$i].SkuId ) -ne $True) { $licensed=$true } } ; If( $licensed -eq $false) { Write-Host $_.UserPrincipalName} }