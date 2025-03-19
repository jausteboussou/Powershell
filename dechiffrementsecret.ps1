$certThumbprint = "***"

# Récupérer le certificat à partir du magasin de certificats de la machine
$cert = Get-Item "Cert:\LocalMachine\My\$certThumbprint"
$encryptedBytes = get-content -path "***"
# Déchiffrer le secret avec le bon type de padding
$rsa = $cert.PrivateKey
$decryptedBytes = $rsa.Decrypt($encryptedBytes, [System.Security.Cryptography.RSAEncryptionPadding]::OaepSHA256)

# Convertir les bytes déchiffrés en une chaîne
$decryptedSecret = [System.Text.Encoding]::UTF8.GetString($decryptedBytes)
Write-Output "Secret déchiffré: $decryptedSecret"
