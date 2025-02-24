# Chargement de l'assembly Windows Forms
Add-Type -AssemblyName System.Windows.Forms

function GraphiqueArchivage {
    # Création du formulaire
    $Form = New-Object System.Windows.Forms.Form
    $Form.Width = 500
    $Form.Height = 150
    $Form.AutoSize = $true
    $Form.Text = "Archivage Utilisateur"

    # Label : Sélectionner un utilisateur
    $LabelSelectUser = New-Object System.Windows.Forms.Label
    $LabelSelectUser.Text = "Selectionner un utilisateur"
    $LabelSelectUser.Location = New-Object System.Drawing.Point(10,12)
    $LabelSelectUser.AutoSize = $true
    $LabelSelectUser.Visible = $true
    $Form.Controls.Add($LabelSelectUser)

    # Liste déroulante pour choisir un utilisateur AD
    $ComboBoxUserList = New-Object System.Windows.Forms.ComboBox
    $ComboBoxUserList.Width = 150
    $ComboBoxUserList.Visible = $true
    $ComboBoxUserList.Location  = New-Object System.Drawing.Point(150,10)
    $Form.Controls.Add($ComboBoxUserList)

    # Bouton pour exécuter le script d'archivage
    $ButtonExecuted = New-Object System.Windows.Forms.Button
    $ButtonExecuted.Location = New-Object System.Drawing.Size(10,80)
    $ButtonExecuted.Size = New-Object System.Drawing.Size(100,25)
    $ButtonExecuted.Text = "Executer"
    $ButtonExecuted.Visible = $true
    $Form.Controls.Add($ButtonExecuted)

    # Ajouter des valeurs à la liste déroulante depuis Active Directory
    Get-ADUser -Filter {Enabled -eq $true} | 
    Select-Object samAccountName | 
    Sort-Object samAccountName | 
    ForEach-Object {
        $ComboBoxUserList.Items.Add($_.SamAccountName)
    }
    

    $ButtonExecuted.Add_Click({
        $inputUser = $ComboBoxUserList.Text.Trim()
    
        if ($ComboBoxUserList.SelectedItem) {
            # Si un utilisateur est sélectionné, on l'affecte
            $script:SelectedUser = $ComboBoxUserList.SelectedItem.ToString()
        }
        elseif ($inputUser -ne "" -and !$ComboBoxUserList.Items.Contains($inputUser)) {
            # Si l'utilisateur est entré manuellement et n'existe pas dans la liste, on l'ajoute
            $ComboBoxUserList.Items.Add($inputUser)
            $script:SelectedUser = $inputUser
        }
        else {
            # Ferme le formulaire sans sélection
            $Form.Close()
            exit
        }
    
        # Ferme le formulaire après l'ajout
        $Form.Close()
    })
    

    $Form.ShowDialog()
if ($script:SelectedUser) {
    $result = [System.Windows.Forms.MessageBox]::Show("Utilisateur sélectionné : $script:SelectedUser. Voulez-vous continuer ?", "Confirmation", 4)  # 4 = YesNo

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        [System.Windows.Forms.MessageBox]::Show("Action validée pour l'utilisateur : $script:SelectedUser")
        # sort
    } else {
        [System.Windows.Forms.MessageBox]::Show("Action annulée.")
        exit
    }
} else {
    [System.Windows.Forms.MessageBox]::Show("Aucun utilisateur sélectionné. Fin du script")
    exit
}

    # Affichage de la sélection après la fermeture du formulaire
    Write-Output "Utilisateur sélectionné : $script:SelectedUser"
}

 
$thumbprintMapping = @{
    'SRV' = '***';
}
 
#Connexion sur le tenant
function ConnectGraphApplication {
    $ClientId = '***'
    $TenantId = '***'
 
    # Get the current server name
    $ServerName = $env:COMPUTERNAME
    Write-Host "Current server: $ServerName" -ForegroundColor Cyan
 
    # Retrieve the thumbprint for the current server
    if ($thumbprintMapping.ContainsKey($ServerName)) {
        $CertThumbprint = $thumbprintMapping[$ServerName]
        Write-Host "Using thumbprint: $CertThumbprint for server: $ServerName" -ForegroundColor Green
    } else {
        Write-Host "No thumbprint found for server: $ServerName. Exiting..." -ForegroundColor Red
        return
    }
 
    # Get the certificate
    $Cert = Get-ChildItem Cert:\LocalMachine\My\$CertThumbprint
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
 
    # Connect to Microsoft Graph
    try {
        Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -Certificate $Cert
        Write-Host "Successfully connected to Microsoft Graph with thumbprint: $CertThumbprint" -ForegroundColor Green
    } catch {
        Write-Host "Failed to connect to Microsoft Graph: $_" -ForegroundColor Red
    }
}
 
 
 
# Fonction d'installation du module ExchangeOnlineManagement
 
 
function Verify-And-ImportModules {
    [CmdletBinding()]
    param (
        [string[]]$Modules = @(
            'ExchangeOnlineManagement',
            'Microsoft.Graph.Applications',
            'Microsoft.Graph.Users',
            'Microsoft.Graph.Authentication'
            'Microsoft.Graph.Users.Actions'
            'Microsoft.Graph.DirectoryObjects'
            'Microsoft.Graph.Identity.DirectoryManagement'
        )
    )
 
    foreach ($Module in $Modules) {
        Write-Host "Checking if module '$Module' is installed..." -ForegroundColor Yellow
 
        if (-not (Get-Module -ListAvailable -Name $Module)) {
            Write-Host "Module '$Module' is not installed. Installing..." -ForegroundColor Red
 
            try {
                Install-Module -Name $Module -Force -Scope CurrentUser -AllowClobber
                Write-Host "Module '$Module' installed successfully." -ForegroundColor Green
            } catch {
                Write-Host "Failed to install module '$Module': $_" -ForegroundColor Red
                continue
            }
        } else {
            Write-Host "Module '$Module' is already installed." -ForegroundColor Green
        }
 
        Write-Host "Importing module '$Module'..." -ForegroundColor Yellow
        try {
            Import-Module -Name $Module -Force
            Write-Host "Module '$Module' imported successfully." -ForegroundColor Green
        } catch {
            Write-Host "Failed to import module '$Module': $_" -ForegroundColor Red
        }
    }
 
    Write-Host "All modules checked and processed." -ForegroundColor Cyan
}
 
# Example usage:
 
 
# Connexion à ExchangeOnline et MsGraph
function Connect-Services {
    Connect-ExchangeOnline
    ConnectGraphApplication
}
 
# Désactivation d'un compte utilisateur
 
function Disable-UserAccount {
    param ($username)
    try {
        if ($username.Enabled -eq $true)
        {
            Write-Host -BackgroundColor Green "Compte en cours de désactivation..."
            Disable-ADAccount -Identity $username
            Write-Host "Le compte utilisateur $username a été désactivé avec succès."
           
        }
        else {
            Write-Host -BackgroundColor Green "Le compte utilisateur $username est déjà désactivé."
        }
    }
    catch {
        Write-Host " ERREUR: desactivation du compte a echoue"
    }
    }
 
# Vérification du statut actuel du compte
 
function CheckUsersStatus {
    param ($username)
    $userInfo = $username | Select-Object Name, Enabled
    Write-Host "Statut actuel du compte de $($userInfo.Name) : $($userInfo.Enabled)"
}
 
# Retirer l'utilisateur de tous les groupes sauf "Domain Users"
 
function Remove-FromGroups {
    param ($username)
    $groups = Get-ADPrincipalGroupMembership -Identity $username.SamAccountName
    try {
        foreach ($group in $groups) {
            if ($group.Name -ne "Domain Users") {
               
                    Remove-ADPrincipalGroupMembership -Identity $username.SamAccountName -MemberOf $group -Confirm:$false
                    Write-Host "L'utilisateur $username a été retiré du groupe $($group.Name)."
            }
        }
    }
    catch {
        Write-Host "ERREUR: lors de la supression du groupe suivant: $($group.Name): $($_.Exception.Message)"
    }
}
 
# Suppression du manager de l'utilisateur
 
function Clear-UserManager {
    param ($username)
    try {
        $removemanager = $username
        if ($removemanager.Manager) {
            Set-ADUser -Identity $username -Clear Manager
            Write-Host "Le manager a été retiré avec succès pour l'utilisateur $username."
        }
        else {
              Write-Host -BackgroundColor Green "Aucun manager est defini pour l'utilisateur $username."
        }
         
    } catch {
        Write-Host -ForegroundColor Red "ERREUR: lors de la suppression du manager pour l'utilisateur $username : $($_.Exception.Message)"
    }
}
 
# Suppression de toutes les licences assignées à l'utilisateur
 

 
# Déplacement de l'utilisateur dans une OU spécifique pour les comptes désactivés
 
function Move-UserToDisabledOU {
    param ($username)
    try {
        Move-ADObject -Identity $username.DistinguishedName   -TargetPath "OU=***,OU=Utilisateurs,OU=***,DC=***,DC=com"
    }
    catch {
        Write-Host "ERREUR: lors du deplacement du user dans OU desactivation."
    }
}
 
function CheckSynchroUser {
    # création du fichier de synchronisation
    Write-Host "Lancement de la synchronisation"
    New-Item -ItemType File -Path "\\***" -Force
    # Check pour la synchronisation du serveur Azure Ad Sync
   
    while (Test-Path -Path "\\***") {
        # doit checker si le fichier est présent
    }
     Start-Sleep 120 # attendre 3 minutes apres synchro
}
function deleteUserEntra{
    param ($username)
    Remove-MgUser -UserId $username.UserPrincipalName -ErrorAction Ignore
}
function Restore-ArchivedUser {
    param ($username)
    try {
            $TodayDate = Get-Date -Format "yyyyMMdd"
            $restoreuser = "ARCHIVE_$($TodayDate)_$($username.SamAccountName)"
            # Récupération de l'objet supprimé correspondant à l'utilisateur
            $deletedUser =  Get-MgDirectoryDeletedUser -All | Where-Object {$_.DisplayName -eq $username.DisplayName }
            # Restauration de l'utilisateur supprimé
            Restore-MgDirectoryDeletedItem -DirectoryObjectId $deletedUser.Id
            # Mise à jour du UserPrincipalName avec un nouveau nom basé sur l'archivage
            $domain = $username.UserPrincipalName.Split("@")[1]
            $newUserPrincipalName = "$restoreuser@$domain"
            # Met à jour l'utilisateur restauré avec le nouveau UserPrincipalName
            Update-MgUser -UserId $username.UserPrincipalName -UserPrincipalName $newUserPrincipalName
            Write-Host "L'utilisateur a été restauré avec succès avec le nouvel UPN : $newUserPrincipalName"
            $newUserPrincipalName = get-mguser -all | Where-Object {$_.UserPrincipalName -eq $newUserPrincipalName}
    }
    catch {
        Write-Host "ERREUR: lors de la restauration de l'utilisateur"
    }
    return $newUserPrincipalName
}
function Update-DisplayName {
    param ($username, $newUserPrincipalName)
    $TodayDate = Get-Date -Format "yyyyMMdd"
    $newDisplayName = "ARCHIVE $($TodayDate) $($newUserPrincipalName.displayName)"
    Update-MgUser -UserId $newUserPrincipalName.UserPrincipalName -DisplayName $newDisplayName
}
 
# Affectation de la licence à l'utilisateur archivé
 
function AttributeLicensetoArchivedUser {
    param($newUserPrincipalName)
    Set-MgUserLicense -UserId $newUserPrincipalName.UserPrincipalName -AddLicenses @{SkuId = "***"} -RemoveLicenses @()
}
 
# Conversion de la boîte mail en boîte partagée et suppression de la licence
function Convert-MailboxtoShared {
    param($newUserPrincipalName)
    try {
            $check_shared = $null
            while ($null -eq $check_shared) {
                Start-Sleep -Seconds 140 # attendre deux minutes avant de tenter de convertir en boîte partagée
                Set-Mailbox -Identity $newUserPrincipalName.UserPrincipalName -Type Shared -ErrorAction SilentlyContinue -WarningAction Ignore
                if(get-mailbox -identity $newUserPrincipalName.UserPrincipalName -ErrorAction SilentlyContinue -WarningAction Ignore)
                {
                    $check_shared = (get-mailbox -identity "$($newUserPrincipalName.UserPrincipalName)").IsShared
                    break
                }
            }
        start-sleep -Seconds 10
        Write-Host -BackgroundColor Green "conversion OK"
        start-sleep -s 180
        Set-MgUserLicense -UserId "$($newUserPrincipalName.UserPrincipalName)" -RemoveLicenses (Get-MgUserLicenseDetail -UserId "$($newUserPrincipalName.UserPrincipalName)").SkuId -AddLicenses @()
        start-sleep -s 180
        Update-MgUser -UserId "$($newUserPrincipalName.UserPrincipalName)" -AccountEnabled:$false
}
    catch {
        Write-Host -ForegroundColor Red "ERREUR: lors de la conversion de la boîte mail en boîte partagée : $($_.Exception.Message)"
    }
}
function Update-EmployeeType{
    param ($newUserPrincipalName)
    update-mguser -UserId "$($newUserPrincipalName.UserPrincipalName)" -EmployeeType "archive"
    Write-Host -ForegroundColor Green "Employetype renomme en <archive>."
}
# Fonction principale orchestrant les étapes
 
function Main {
    write-host -ForegroundColor Green "Etape 1 installation des modules"
    # installation des modules
    Verify-And-ImportModules
    write-host -ForegroundColor Green "Etape 2 connexion aux services"
    Connect-Services
    # Entrer le displayname de l'utilisateur
    write-host -ForegroundColor Green "Etape 3 Entrer un utilisateur"
    GraphiqueArchivage | Out-Null
    # obtenir toutes les propriétés de l'utilisateur
    $username = Get-ADUser -Identity $script:SelectedUser -Properties *
    # Voir si l'utilisateur est activé ou désactivé
    write-host -ForegroundColor Green "Etape 4 vérification de l'état de l'utilisateur"
    CheckUsersStatus -username $username
    # Désactivation de l'utilisateur
    write-host -ForegroundColor Green "Etape 5 désactivation du compte AD de l'utilisateur"
    Disable-UserAccount -username $username
    # Suppression de tous les groupes de l'utilisateur
    write-host -ForegroundColor Green "Etape 6 suppression de tous les groupes, excepté domain users"
    Remove-FromGroups -username $username
    # Suppression du manager
    write-host -ForegroundColor Green "Etape 7 suppression du manager de l'utilisateur"
    Clear-UserManager -username $username
    # Déplacement de l'utilisateur dans l'OU Disabled
    write-host -ForegroundColor Green "Etape 8 déplacement de l'utilisateur dans l'OU Disabled"
    Move-UserToDisabledOU -username $username
    # Restauration de l'utilisateur archivé
    write-host -ForegroundColor Green "Etape 9 restauration de l'utilisateur supprimé, tout en modifiant son displayname"
    CheckSynchroUser # synchronisation
    $newUserPrincipalName = Restore-ArchivedUser -username $username
    # Modification du nom de l'utilisateur, ex (Archive_05092023_jboussou)
    # Assignation de la licence E1
    write-host "Etape 10 attribution de la licence E1"
    AttributeLicensetoArchivedUser -newUserPrincipalName $newUserPrincipalName
    # Conversion de la boîte mail en boîte partagée
    write-host "Etape 11 conversion de la boîte mail en boîte partagée"
    Convert-MailboxtoShared -newUserPrincipalName $newUserPrincipalName
    # Modification du nom affiché pour l'utilisateur
    write-host "Etape 12 modification du nom affiché"
    Update-DisplayName -newUserPrincipalName $newUserPrincipalName
    Update-EmployeeType -newUserPrincipalName $newUserPrincipalName
}
# Appel de la fonction principale
Main
