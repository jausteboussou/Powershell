# Fonction pour arrêter les processus FortiClient
function Stop-FortiProcesses {
    Write-Host "Arrêt des processus FortiClient..."

    Set-Service -Name "FA_Scheduler" -StartupType Disabled -ErrorAction SilentlyContinue
    Stop-Process -Name "scheduler" -Force -ErrorAction SilentlyContinue
    # Liste des processus à arrêter
    $processes = @(
        "FortiESNAC",
        "FortiTray",
        "FortiSSLVPNdaemon",
        "FortiSettings",
        "FSSOMA",
        "FCDBLog",
        "FMService64",
        "FortiVPN"
    )

    # Arrêt des processus
    foreach ($process in $processes) {
        try {
            # Utilisation de Get-Process pour vérifier si le processus existe avant de l'arrêter
            $processToStop = Get-Process -Name $process -ErrorAction SilentlyContinue
            if ($processToStop) {
                Stop-Process -InputObject $processToStop -Force -ErrorAction SilentlyContinue
                if ($?) { Write-Host "Processus : '$process' arrêté." }
            } else {
                Write-Host "Le processus '$process' n'est pas en cours d'exécution."
            }
        } catch {
            Write-Warning "Erreur lors de l'arrêt du processus '$process' : $($_.Exception.Message)"
        }
    }
    Write-Host "Arrêt des processus FortiClient terminé."
}

# Fonction pour supprimer les clés de registre
function Remove-RegistryKeys {
    Write-Host "Suppression des clés de registre FortiClient..."

    # Définir les chemins de base à vérifier.
    $basePaths = @(
        "HKLM:\SOFTWARE\Classes\Installer\Products",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Classes\Installer\Products",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\Fortinet",
        "HKLM:\SOFTWARE\WOW6432Node\Fortinet",
        "HKCU:\SOFTWARE\Fortinet",
        "HKLM:\SOFTWARE\fctlog"
    )

    $fortiProductCode = "{6COA3C5E-7725-49D8-A016-B3ADCACF61C2}"
    $registryKeysToRemove = @()

    # Recherche dans les chemins de base
    foreach ($basePath in $basePaths) {
        Write-Host "Recherche dans : '$basePath'"
        try {
            $keys = Get-ChildItem -Path $basePath -ErrorAction SilentlyContinue
            foreach ($key in $keys) {
                try {
                    # Vérifier si la clé correspond exactement au Product Code
                    if ($key.Name -ceq $fortiProductCode) {
                        $registryKeysToRemove += "Registry::$($key.PSPath)"
                        Write-Host "[REGISTRE] : Clé de produit FortiClient trouvée : '$($key.PSPath)'"
                    }
                    # Rechercher des valeurs contenant le nom FortiClient dans les clés Uninstall
                    elseif ($basePath -like "*Uninstall*") {
                        $uninstallInfo = Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue
                        if ($uninstallInfo.DisplayName -like "*FortiClient*" -or $uninstallInfo.Publisher -like "*Fortinet*") {
                            $registryKeysToRemove += "Registry::$($key.PSPath)"
                            Write-Host "[REGISTRE] : Clé de désinstallation FortiClient trouvée : '$($key.PSPath)'"
                        }
                    }
                    # Rechercher des clés ou des valeurs contenant "Forti" dans les autres chemins
                    else {
                        if ($key.Name -like "*Forti*" -or $key.Name -like "*fctlog*") {
                            $registryKeysToRemove += "Registry::$($key.PSPath)"
                            Write-Host "[REGISTRE] : Clé contenant 'Forti' trouvée : '$($key.PSPath)'"
                        }
                        $values = Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue
                        if ($values) {
                            foreach ($propertyName in $values.PSObject.Properties.Name) {
                                if ($values.$propertyName -like "*Forti*") {
                                    $registryKeysToRemove += "Registry::$($key.PSPath)"
                                    Write-Host "[REGISTRE] : Valeur contenant 'Forti' trouvée dans la clé : '$($key.PSPath)'"
                                    break # Une correspondance suffit pour ajouter la clé
                                }
                            }
                        }
                    }
                } catch {
                    Write-Warning "[ERREUR] : Erreur lors de la lecture de la clé '$($key.PSPath)' : $($_.Exception.Message)"
                }
            }
        } catch {
            Write-Warning "[ERREUR] : Erreur lors de l'accès au chemin '$basePath' : $($_.Exception.Message)"
        }
    }

    # Supprimer les clés de registre trouvées (en évitant les doublons)
    $uniqueKeysToRemove = $registryKeysToRemove | Sort-Object -Unique
    if ($uniqueKeysToRemove) {
        Write-Host "Check condition de suppression des clés"
        foreach ($key in $uniqueKeysToRemove) {
            try {
                $splitPath = $key -split '::'
                $keyToRemove = $splitPath[2]
                Remove-Item -Path "Registry::$($keyToRemove)" -Force -Recurse -ErrorAction SilentlyContinue
                if ($?) { Write-Host "[REGISTRE] : Clé '$keyToRemove' supprimée." }
            } catch {
                Write-Warning "[ERREUR] : $($_.Exception.Message) lors de la suppression de la clé de registre '$keyToRemove'"
            }
        }
    } else {
        Write-Host "Aucune clé de registre FortiClient spécifique trouvée."
    }
    Write-Host "Suppression des clés de registre FortiClient terminée."
}

# Fonction pour supprimer les répertoires et fichiers
function Remove-FortiFiles {
    Write-Host "Suppression des fichiers et dossiers FortiClient..."

    $filesAndFolders = @(
        "C:\Program Files\Fortinet\",
        "C:\Users\Public\Desktop\FortiClient.lnk",
        "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\FortiClient"
    )

    foreach ($item in $filesAndFolders) {
        if (Test-Path -Path $item) { # Vérifier si le chemin existe
            try {
                Remove-Item -Path $item -Force -Recurse -ErrorAction SilentlyContinue
                if ($?) { Write-Host "Suppression de : '$item'" }
            } catch {
                Write-Warning "Erreur lors de la suppression de '$item' : $($_.Exception.Message)"
            }
        } else {
            Write-Host "'$item' n'existe pas et n'a pas besoin d'être supprimé."
        }
    }
    Write-Host "Suppression des fichiers et dossiers FortiClient terminée."
}

# Exécution du script
Write-Host "--------------------------"
Write-Host "Début de la suppression de FortiClient"
Write-Host "--------------------------"

# Tenter de réinitialiser le mot de passe (peut ne pas toujours fonctionner)
Set-ItemProperty -Path "HKLM:\SOFTWARE\Fortinet\FortiClient\FA_FCM" -Name "password" -Value "" -ErrorAction SilentlyContinue
if ($?) { Write-Host "Tentative de réinitialisation du mot de passe FortiClient." }

Stop-FortiProcesses
Write-Host "--------------------------"
Remove-RegistryKeys
Remove-FortiFiles
Write-Host "--------------------------"
Write-Host "Suppression de FortiClient terminée"
Write-Host "--------------------------"
