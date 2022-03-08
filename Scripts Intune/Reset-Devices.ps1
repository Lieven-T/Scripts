Connect-MSGraph
Connect-AzureAD
Get-AzureADGroup -Filter "startswith(displayName,'romerocollege_2')" -All $true | % {
    $Title = $_.DisplayName -replace "romerocollege_",""
    $Delimiter = "=" * $Title.Length
    Write-Host "`n$Delimiter"
    Write-Host $Title
    Write-Host $Delimiter
    Get-AzureADGroupMember -ObjectId $_.ObjectId | % {
        Write-Host $_.DisplayName
        (Invoke-MSGraphRequest -Url "https://graph.microsoft.com/Beta/users/$($_.ObjectId)/managedDevices").value | % {
            Write-Host "    $($_.deviceName)"
            # Invoke-MSGraphRequest -HttpMethod POST -Url "https://graph.microsoft.com/Beta/deviceManagement/managedDevices/$($_.id)/wipe" -Content "{keepEnrollmentData:`"true`", keepUserData:`"false`"}"
        }
    }
}
