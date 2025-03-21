<#
--------------------------- DESCRIPTION DU SCRIPT ----------------------------

Je vais expliquer en détail le but et le fonctionnement de ce script :

Objectif Principal :
Ce script extrait et exporte les informations des comptes utilisateurs archivés dans Microsoft 365 en se basant sur une convention de nommage spécifique.
Contexte et Problématique

Dans Microsoft Entra (anciennement Azure AD), lors de l'archivage d'un compte utilisateur, la seule trace de la date d'archivage se trouve dans le UserPrincipalName qui suit le format :

ARCHIVE_AAAAMMJJ_utilisateur@domaine.fr
je prend donc l'index 1 en me basant dans le code de la fonction split.

Exemple : ARCHIVE_20231215_jboussou@domaine.fr
1. Les comptes "archives"
2. Les comptes "redirection"

Étapes Détaillées :

1. Préparation :
- Vérifie si le module ImportExcel est installé
- Se connecte à Microsoft Graph (pour accéder aux données Microsoft 365)
- Importe les modules nécessaires

2. Pour les Comptes Archives :
- Recherche tous les utilisateurs dont le nom contient "archive"
- Extrait la date d'archivage du nom d'utilisateur
- Collecte les informations : 
  * Nom d'utilisateur principal (UserPrincipalName)
  * Nom d'affichage (DisplayName)
  * Nom de famille (Surname)
  * Date d'archivage

3. Pour les Comptes Redirection :
- Recherche tous les utilisateurs dont le nom contient "redirection"
- Collecte les informations :
  * Nom d'utilisateur principal
  * Nom d'affichage
  * Date de création

4. Export des Données :
- Crée un fichier Excel avec la date du jour
- Crée deux onglets distincts :
  * Un pour les comptes archives
  * Un pour les comptes redirection
- Formate les données en tableaux Excel automatiquement dimensionnés

5. Gestion des Erreurs :
- Inclut des mécanismes de gestion d'erreurs à plusieurs niveaux
- Affiche des messages d'erreur explicites
- Gère les exceptions potentielles

6. Retours Utilisateur :
- Affiche le nombre de comptes trouvés dans chaque catégorie
- Confirme la fin de l'export
- Indique le chemin du fichier créé

Utilité du Script :
- Permet un suivi des comptes archives et redirection
- Facilite l'audit des comptes
- Automatise la création de rapports
- Aide à la gestion et au nettoyage des comptes utilisateurs

Le script est particulièrement utile pour :
- Les administrateurs IT
- Les équipes de sécurité
- La gestion des accès
- L'audit des comptes utilisateurs

Le résultat final est un fichier Excel bien organisé, facile à lire et à analyser, mis à jour quotidiennement avec les dernières informations des comptes.

#>

try 
{
    # Vérifie si le module ImportExcel est installé, sinon l'installe
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "Installation du module ImportExcel..."
        Install-Module -Name ImportExcel -Force -Scope CurrentUser
    }

    # Fonction pour récupérer les utilisateurs avec "archive" dans leur nom
    function Get-AllArchiveUsers {
        $ExportAllArchiveUsers = @() # Initialise un tableau vide
    
        # Récupère tous les utilisateurs "archive" et ajoute une colonne ArchiveDate
        $AllArchiveUsers = Get-MgUser -All -Property UserPrincipalName, DisplayName, Surname, createdDateTime | 
            Where-Object { $_.UserPrincipalName -match "archive" -or $_.DisplayName -match "archive" } |
            Select-Object *, @{
                Name = 'ArchiveDate';
                Expression = {
                    try {
                        # Extrait la date du UserPrincipalName et la formate
                        $matche = ($_.UserPrincipalName -split '_')[1]
                        $date = [DateTime]::ParseExact($matche, "yyyyMMdd", $null)
                        return $date.ToString("dd/MM/yyyy")
                    }
                    catch {
                        $null
                    }
                }
            }
    
        # Crée un objet personnalisé pour chaque utilisateur
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
    
    # Fonction pour récupérer les utilisateurs avec "redirection" dans leur nom
    function Get-AllRedirectionUsers {
        $ExportAllRedirectionUsers = @() # Initialise un tableau vide
        # Récupère tous les utilisateurs "redirection"
        $AllRedirectionUsers = Get-MgUser -All -Property UserPrincipalName, DisplayName,createdDateTime | 
        Select-Object UserPrincipalName, DisplayName, createdDateTime | 
        Where-Object {($_.UserPrincipalName -match "redirection" -or $_.DisplayName -match "redirection")}
        
        # Crée un objet personnalisé pour chaque utilisateur
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

    # Fonction pour se connecter à Microsoft Graph
    function connectgraph {
        connect-mggraph -Scopes "User.Read.All"
    }

    # Fonction pour importer les modules nécessaires
    function importmodule {
        import-module -name "Microsoft.Graph.Users"
        Import-Module -Name ImportExcel
    }

    # Fonction principale qui orchestre tout le processus
    function main {
        $Today = Get-Date -UFormat "%Y-%m-%d" # Obtient la date du jour
        connectgraph # Connexion à Graph
        importmodule # Import des modules
        
        # Définit le chemin du fichier Excel
        $ExportPath = "D:\***\***5\***$($Today).xlsx"

        # Récupère les données des deux types d'utilisateurs
        $ArchiveUsers = Get-AllArchiveUsers
        $RedirectionUsers = Get-AllRedirectionUsers

        # Export vers Excel dans deux onglets différents
        $ArchiveUsers | Export-Excel -Path $ExportPath -WorksheetName 'Archives' -AutoSize -TableName 'Archives'
        $RedirectionUsers | Export-Excel -Path $ExportPath -WorksheetName 'Redirections' -AutoSize -TableName 'Redirections'

        Write-Host -ForegroundColor Green "Export terminé vers $ExportPath"
    }

    # Exécution de la fonction principale
    main
}
catch 
{
    # Gestion des erreurs globales
    Write-Host -ForegroundColor Red "[ERREUR] : Une erreur lors de l'appel des fonctions : $($_.Exception)"
}
