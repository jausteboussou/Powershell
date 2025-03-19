# Emplacement du certificat auto-sign� dans le magasin de certificats
$certThumbprint = "***"

# R�cup�rer le certificat � partir du magasin de certificats de la machine
$cert = Get-Item "Cert:\LocalMachine\My\$certThumbprint"

# Secret � chiffrer
$secret = "***"

# Convertir le secret en bytes
$secretBytes = [System.Text.Encoding]::UTF8.GetBytes($secret)

# Chiffrer le secret avec le bon type de padding
$rsaProvider = [System.Security.Cryptography.RSACryptoServiceProvider]::new()
$rsaProvider.FromXmlString($cert.PublicKey.Key.ToXmlString($false))
$encryptedBytes = $rsaProvider.Encrypt($secretBytes, [System.Security.Cryptography.RSAEncryptionPadding]::OaepSHA1)

# Convertir les bytes chiffr�s en une cha�ne Base64 pour l'affichage
$encryptedSecret = [Convert]::ToBase64String($encryptedBytes)

# Afficher le secret chiffr�
Write-Output "Secret chiffré: $encryptedSecret"
