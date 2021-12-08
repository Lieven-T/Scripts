$RootLocation = "D:\ll\"
$Class = "7OA"
Get-ChildItem "$RootLocation$Class" | % { 
    Write-Host "WinFakt kopiëren naar $_..."
    Copy-Item D:\util\software\WinFakt\WinFakt-Scholen $_.FullName -Recurse 
}