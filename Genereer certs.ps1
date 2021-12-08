$cert = New-SelfSignedCertificate -Type Custom -KeySpec Signature `
-Subject “CN=azureorc” -KeyExportPolicy Exportable `
-HashAlgorithm sha256 -KeyLength 2048 `
-CertStoreLocation “Cert:\CurrentUser\My” -KeyUsageProperty Sign -KeyUsage CertSign


Get-ChildItem -Path “Cert:\CurrentUser\My”
$cert = Get-ChildItem -Path “Cert:\CurrentUser\My\A360D9C3EFCCD0DFC17ADBEE8E86E2DEDB0DA43A”

New-SelfSignedCertificate -Type Custom -KeySpec Signature `
-Subject “CN=azureorcclient” -KeyExportPolicy Exportable -NotAfter (Get-Date).AddYears(1) `
-HashAlgorithm sha256 -KeyLength 2048 `
-CertStoreLocation “Cert:\CurrentUser\My” `
-Signer $cert -TextExtension @(“2.5.29.37={text}1.3.6.1.5.5.7.3.2”)