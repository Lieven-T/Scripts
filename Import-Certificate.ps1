using namespace System.Security.Cryptography.X509Certificates
$store = [X509Store]::new('My', 'CurrentUser', 'ReadWrite')
$store.Add([X509Certificate2]::new('/home/rspletzer/mycert.pfx', '<passphrase>', [X509KeyStorageFlags]::PersistKeySet))
$store.Dispose()