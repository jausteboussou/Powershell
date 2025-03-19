# Emplacement du certificat auto-sign� dans le magasin de certificats
$certThumbprint = "***"

# R�cup�rer le certificat � partir du magasin de certificats de la machine
$cert = Get-Item "Cert:\LocalMachine\My\$certThumbprint"

# Secret � chiffrer
$secret = "***"

# Convertir le secret en bytes
$secretBytes = [System.Text.Encoding]::UTF8.GetBytes($secret)

# Chiffrer le secret avec le bon type de padding
$rsa = [System.Security.Cryptography.RSA]::Create()
$rsa.FromXmlString($cert.PublicKey.Key.ToXmlString($false))
$encryptedSecret = $rsa.Encrypt($secretBytes, [System.Security.Cryptography.RSAEncryptionPadding]::OaepSHA256)


# Afficher le secret chiffr�
Write-Output "Secret chiffré: $encryptedSecret"

