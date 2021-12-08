try {
    Get-AzureADTenantDetail | Out-Null
} catch {
    $Credentials = Get-Credential

    Connect-AzureAD -ErrorAction Stop -Credential $Credentials
    Connect-MicrosoftTeams -Credential $Credentials

    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    Import-PSSession $Session -DisableNameChecking
}
