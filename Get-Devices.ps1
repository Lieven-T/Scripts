Connect-AzureAD
Get-AzureADDevice | ? DeviceTrustType -eq "AzureAD" | % {
    $User = $_ | Get-AzureADDeviceRegisteredOwner

    $Props = @{
        Device = $_.DisplayName
        Enabled = $_.AccountEnabled
        User = $User.UserPrincipalName
        Role = $User.JobTitle
        Department = $User.Department
    }
    New-Object -TypeName psobject -Property $Props
} | Export-Excel C:\temp\laptops.xlsx