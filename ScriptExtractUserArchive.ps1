try 
{
    # Vérification et installation du module ImportExcel si nécessaire
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "Installation du module ImportExcel..."
        Install-Module -Name ImportExcel -Force -Scope CurrentUser
    }

    function Get-AllArchiveUsers {
        $ExportAllArchiveUsers = @()
    
        $AllArchiveUsers = Get-MgUser -All -Property UserPrincipalName, DisplayName, Surname, createdDateTime | 
            Where-Object { $_.UserPrincipalName -match "archive" -or $_.DisplayName -match "archive" } |
            Select-Object *, @{
                Name = 'ArchiveDate';
                Expression = {
                    try {
                        $matche = ($_.UserPrincipalName -split '_')[1]
                        $date = [DateTime]::ParseExact($matche, "yyyyMMdd", $null)
                        return $date.ToString("dd/MM/yyyy")
                    }
                    catch {
                        $null
                    }
                }
            }
    
        foreach ($User in $AllArchiveUsers) {
            $ExportArchiveUser = [PSCustomObject]@{
                UserPrincipalName = $User.UserPrincipalName
                DisplayName = $User.DisplayName
                Surname = $User.Surname
                ArchiveDate = $User.ArchiveDate
            }
            $ExportAllArchiveUsers += $ExportArchiveUser
        }
    
        Write-Host -ForegroundColor Green "Export de $($ExportAllArchiveUsers.Count) comptes archives"
        return $ExportAllArchiveUsers
    }
    
    function Get-AllRedirectionUsers {
        $ExportAllRedirectionUsers = @()
        $AllRedirectionUsers = Get-MgUser -All -Property UserPrincipalName, DisplayName,createdDateTime | 
        Select-Object UserPrincipalName, DisplayName, createdDateTime | 
        Where-Object {($_.UserPrincipalName -match "redirection" -or $_.DisplayName -match "redirection")}
        
        foreach ($AllRedirectionUser in $AllRedirectionUsers) {
            $ExportRedirectionUser = [PSCustomObject]@{    
                UserPrincipalName = $AllRedirectionUser.UserPrincipalName
                DisplayName = $AllRedirectionUser.DisplayName
                CreatedDateTime = try{
                    ($AllRedirectionUser.createdDateTime).ToString("dd/MM/yyyy")
                } catch{Write-Host "[ERREUR] : conversion format"}
            }
            $ExportAllRedirectionUsers += $ExportRedirectionUser    
        }
        Write-Host -ForegroundColor Green "Export de $($ExportAllRedirectionUsers.Count) comptes redirection"
        return $ExportAllRedirectionUsers
    }

    function connectgraph {
        connect-mggraph -Scopes "User.Read.All"
    }

    function importmodule {
        import-module -name "Microsoft.Graph.Users"
        Import-Module -Name ImportExcel
    }

    function main {
        $Today = Get-Date -UFormat "%Y-%m-%d"
        connectgraph
        importmodule
        
        # Changement de l'extension de .csv à .xlsx
        $ExportPath = "D:\***\***\***$($Today).xlsx"

        # Récupération des données
        $ArchiveUsers = Get-AllArchiveUsers
        $RedirectionUsers = Get-AllRedirectionUsers

        # Export vers Excel avec deux onglets distincts
        $ArchiveUsers | Export-Excel -Path $ExportPath -WorksheetName 'Archives' -AutoSize -TableName 'Archives'
        $RedirectionUsers | Export-Excel -Path $ExportPath -WorksheetName 'Redirections' -AutoSize -TableName 'Redirections'

        Write-Host -ForegroundColor Green "Export terminé vers $ExportPath"
    }

    # Appel du main
    main
}
catch 
{
    Write-Host -ForegroundColor Red "[ERREUR] : Une erreur lors de l'appel des fonctions : $($_.Exception)"
}
