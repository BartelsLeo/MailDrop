<#
.SYNOPSIS
    Vertraut dem selbstsignierten MailDrop-Zertifikat, damit die ClickOnce-Installation
    (setup.exe / MailDrop.vsto) auf diesem Rechner ohne Sicherheitswarnung funktioniert.

.DESCRIPTION
    MailDrop wird mit einem von Visual Studio automatisch erzeugten, selbstsignierten
    Zertifikat signiert (kein Zertifikat einer offiziellen Zertifizierungsstelle).
    Windows/Office vertrauen diesem Aussteller auf einem fremden Rechner deshalb nicht,
    wodurch die Installation mit einer Zertifikatswarnung fehlschlaegt oder abgebrochen wird.

    Dieses Skript importiert das (oeffentliche) MailDrop-Zertifikat in die Zertifikatsspeicher
    "Vertrauenswuerdige Stammzertifizierungsstellen" und "Vertrauenswuerdige Herausgeber"
    des aktuellen Benutzers (CurrentUser-Scope). Es enthaelt keinen privaten Schluessel und kann
    daher gefahrlos veroeffentlicht werden. Es sind KEINE Administratorrechte noetig, da die
    Zertifikatsspeicher des aktuellen Benutzers ohne erhoehte Rechte beschreibbar sind.

.PARAMETER Uninstall
    Entfernt das MailDrop-Zertifikat wieder aus den Zertifikatsspeichern.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Install-Certificate.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Install-Certificate.ps1 -Uninstall
#>

[CmdletBinding()]
param(
    [switch]$Uninstall
)

$ErrorActionPreference = 'Stop'

# Oeffentliches MailDrop-Signaturzertifikat (kein privater Schluessel), extrahiert direkt nach der
# Zertifikatserzeugung. Gueltig bis 01.07.2056 (30 Jahre) - bewusst langlebig gewaehlt, damit dieser
# Trust-Schritt nicht alle paar Monate/Jahre fuer bereits installierte Benutzer wiederholt werden muss.
# Bei einer neuen Zertifikatsgenerierung (z.B. Kompromittierung des privaten Schluessels) muss dieser
# Block aus dem neuen MailDrop.vsto (Element <X509Certificate>) aktualisiert werden.
$certBase64 = @'
MIIC+DCCAeCgAwIBAgIQPb0BPSoZZ6ZIMpoNh+YkOzANBgkqhkiG9w0BAQsFADAT
MREwDwYDVQQDDAhNYWlsRHJvcDAgFw0yNjA3MDExODM0MThaGA8yMDU2MDcwMTE4
NDQxOFowEzERMA8GA1UEAwwITWFpbERyb3AwggEiMA0GCSqGSIb3DQEBAQUAA4IB
DwAwggEKAoIBAQCWwaY+JKeuZQQ6WcOomYXhlS972GHb8Siz23pIhofPL+zoiTX+
gsEW+Jb0qNpV8EhUCuuo5YJzBhabHGf02rCGCWKKklrE80bE5Gf2mbQNoPm48ejy
+QRLXJD3IP4KydwsIl9tm/28Wx1csBaxt+SVBS0NzTeqVe/ACfSOolEPZM+xcl2p
lXTz8Vspg0ZDRixIXuIYZYlTtJ7V33Cml6jIsfhQldNr734tpQ7/AoR/Nhx8Y/V+
dXJmvt7y5JMtHCSUePp8qCfti3DJnwSUL+kFwKQVcxkAQHWg0aQxl+zOWZOvd9RQ
Gd3WzAPZp7Ui+/lAOlqtTJFQiow8hzuDRQ9dAgMBAAGjRjBEMA4GA1UdDwEB/wQE
AwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNVHQ4EFgQUVj10Ue7vEf68lSGf
YU0gsp6HCj4wDQYJKoZIhvcNAQELBQADggEBAGnP077LGrDvNLacv6+HmmUwi+5S
o++ouXUZ7SCSPpycoswoxqNZ8nASiq+K5csl9bq3oY2d3UCOW0zef3tCoZSKHCkr
ZUgkzPmDtUH+I1RrgFp08T2QrHXZ1VkU+Cs/Vjn76sIifHI63kUbAxCYp/0Ka5Fk
Jvl2RQFemU5H0hLKKZUp8fadilEbAcWNHRH0xzRXPFHgqeikv32iAOBqMiF3GVQI
FeMSO+U61FTadAmxys/4P6O3IMC9hGtRR/CaGe3btD8gsVFWXEfvz2PKfAlQ0m6O
RYYKww5LsO+Nt0brzhDRDzqFQLN/+8lB2U2Tyto6mEJclz4Rznd/ob4mzUo=
'@

$storeNames = @('Root', 'TrustedPublisher')

function Get-MailDropCertificate {
    $bytes = [Convert]::FromBase64String($certBase64)
    return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(, $bytes)
}

$cert = Get-MailDropCertificate
Write-Host "MailDrop-Zertifikat:" -ForegroundColor Cyan
Write-Host "  Aussteller : $($cert.Subject)"
Write-Host "  Thumbprint : $($cert.Thumbprint)"
Write-Host "  Gueltig bis: $($cert.NotAfter)"
Write-Host ""

if ($cert.NotAfter -lt (Get-Date)) {
    Write-Warning "Dieses Zertifikat ist abgelaufen. Bitte pruefen, ob im Publish-Ordner eine neuere Version dieses Skripts vorliegt."
}

foreach ($storeName in $storeNames) {
    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($storeName, 'CurrentUser')
    $store.Open('ReadWrite')
    try {
        if ($Uninstall) {
            $existing = $store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
            if ($existing) {
                $store.Remove($cert)
                Write-Host "Entfernt aus 'CurrentUser\$storeName'." -ForegroundColor Yellow
            } else {
                Write-Host "War nicht in 'CurrentUser\$storeName' vorhanden." -ForegroundColor DarkGray
            }
        } else {
            $store.Add($cert)
            Write-Host "Erfolgreich hinzugefuegt zu 'CurrentUser\$storeName'." -ForegroundColor Green
        }
    }
    finally {
        $store.Close()
    }
}

Write-Host ""
if ($Uninstall) {
    Write-Host "Fertig. Das MailDrop-Zertifikat wird nicht mehr als vertrauenswuerdig eingestuft." -ForegroundColor Cyan
} else {
    Write-Host "Fertig. Die Installation von MailDrop (setup.exe / MailDrop.vsto) sollte jetzt ohne" -ForegroundColor Cyan
    Write-Host "Zertifikatswarnung funktionieren. Bitte 'setup.exe' aus diesem Ordner erneut ausfuehren." -ForegroundColor Cyan
}
