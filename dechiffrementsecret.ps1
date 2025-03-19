$certThumbprint = "***"

# Récupérer le certificat à partir du magasin de certificats de la machine
$cert = Get-Item "Cert:\LocalMachine\My\$certThumbprint"

# Secret chiffré (en Base64)
$encryptedSecretBase64 = get-content -Path "***"

# Convertir le secret chiffré de Base64 en bytes
$encryptedBytes = [Convert]::FromBase64String($encryptedSecretBase64)

# Accéder à la clé privée RSA directement
$rsaProvider = $cert.PrivateKey

# Déchiffrer le secret
$decryptedBytes = $rsaProvider.Decrypt($encryptedBytes, [System.Security.Cryptography.RSAEncryptionPadding]::OaepSHA1)

# Convertir les bytes déchiffrés en chaîne
$decryptedSecret = [System.Text.Encoding]::UTF8.GetString($decryptedBytes)

# Afficher le secret déchiffré
Write-Output "Secret déchiffré: $decryptedSecret"
