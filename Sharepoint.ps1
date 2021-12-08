<#
Install-Module -Name PnP.PowerShell -RequiredVersion 0.3.9-nightly -AllowPrerelease
Import-Module PnP.Powershell
Register-PnPManagementShellAccess
#>

$SiteUrl = "https://romerocollege.sharepoint.com/sites/"
$InputData = Import-Excel .\teams.xlsx

$Cred = Get-Credential
Connect-AzureAD -Credential $Cred

$InputData | Select -ExpandProperty Klas -Unique | % {
    try {
        $Class = $_
        Write-Host "Verwerken klas $Class"
        Connect-PnPOnline -Url ($SiteUrl + $Class) -Credentials $Cred -ErrorAction Stop
        $Owners = Get-PnPGroup -AssociatedOwnerGroup
        $Users = Get-AzureADGroup -SearchString $Class | Get-AzureADGroupMember | ? UserPrincipalName -Match "@student"

        $InputData | ? Klas -eq $Class | select -ExpandProperty Vak -Unique | % {
            $Vak = $_
            Write-Host "    Verwerken vak $Vak"
            $VakFolder = "Gedeelde documenten/$Vak"
            try {
                Get-PnPFolder $VakFolder -ErrorAction Stop | Out-Null
                $Subfolders = Get-PnPFolderItem -FolderSiteRelativeUrl $VakFolder
                $Users | ? { ($_.UserPrincipalName -split "@")[0] -notin ($Subfolders | select -ExpandProperty Name) } | % {
                    $UserName = ($_.UserPrincipalName -split "@")[0]
                    Write-Host "        Aanmaken map $UserName"
                    $Folder = Add-PnPFolder -Name $UserName -Folder $VakFolder
                    Set-PnPFolderPermission -List $VakFolder -Identity "$VakFolder/$UserName" -User $_.UserPrincipalName -AddRole 'Contribute' -ClearExisting
                    Set-PnPFolderPermission -List $VakFolder -Identity "$VakFolder/$UserName" -Group $Owners -AddRole 'Contribute'
                }

                $SubFolders | ? Name -notin ($Users | % { ($_.UserPrincipalName -split "@")[0] }) | % {
                    Write-Host "        Schrappen map $($_.Name)"
                    Remove-PnPFolder -Name $_.Name -Folder $VakFolder -Force
                }
            } catch [Microsoft.SharePoint.Client.ServerException] {
                Write-Host "        $_"
            }
        }
    } catch {
        Write-Host "    $_"
    }
}