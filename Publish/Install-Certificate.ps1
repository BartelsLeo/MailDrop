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

    Entfernt ausserdem IMMER (auch ohne -Uninstall) das/die alte(n), kompromittierte(n)
    MailDrop-Zertifikat(e) aus $oldCompromisedThumbprints, falls vorhanden - siehe
    "Zertifikatsrotation" unten.

.PARAMETER Uninstall
    Entfernt das aktuelle MailDrop-Zertifikat wieder aus den Zertifikatsspeichern.

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
# Zertifikatserzeugung. Gueltig bis 11.09.2056 (~30 Jahre) - bewusst langlebig gewaehlt, damit dieser
# Trust-Schritt nicht alle paar Monate/Jahre fuer bereits installierte Benutzer wiederholt werden muss.
# Bei einer neuen Zertifikatsgenerierung (z.B. Kompromittierung des privaten Schluessels) muss dieser
# Block aus dem neuen MailDrop.vsto (Element <X509Certificate>) aktualisiert werden - UND der alte
# Thumbprint unten in $oldCompromisedThumbprints eingetragen werden, damit dieses Skript alte,
# kompromittierte Zertifikate auch aktiv wieder aus dem Trust-Store entfernt statt sie nur additiv
# stehen zu lassen.
$certBase64 = @'
MIIDDTCCAfWgAwIBAgIUAhc/eJAL6qk23AO0F5984QRiSR8wDQYJKoZIhvcNAQEL
BQAwEzERMA8GA1UEAwwITWFpbERyb3AwIBcNMjYwOTE5MTgzMDE3WhgPMjA1NjA5
MTExODMwMTdaMBMxETAPBgNVBAMMCE1haWxEcm9wMIIBIjANBgkqhkiG9w0BAQEF
AAOCAQ8AMIIBCgKCAQEAqM2wiOcN0+4NofA8yCj7slyIxMujh6yM5188iF3OMiTZ
jv4WfIyvEuNtSmT9RXCcNjAmQjLS7nhwnmjPbOJHKuOvGzcv8oMD3f2TV1CnGCU2
twfn7mz+siC8r2WZ02KMcrnR3o8xOOX/XaRKWZt3guzY3jVbA8LOl0yxteHTmYhw
XHXICUmGepVRLuoJ+9ytJkgFUtABTgVCbFbANUXYAryFUGGN78BBjMm/ZgCJ3O0W
51DXQzUKY1SzBIR1JLxe1MGgl6QfX2/CflADI8cIeYFLlhXNdQQwx3RREyQ3beyZ
q4FOUBVqJHOEs+9vdsZG7m20eLw51HrDWXXKaakv9QIDAQABo1cwVTAMBgNVHRMB
Af8EAjAAMA4GA1UdDwEB/wQEAwIHgDAWBgNVHSUBAf8EDDAKBggrBgEFBQcDAzAd
BgNVHQ4EFgQURI6HdlUwnLqbnWVJEWQEZxh8ZFMwDQYJKoZIhvcNAQELBQADggEB
AFV3ofGmliVFDteL1RxKMpuCGSLPRlvRPQ+X4QVWY8EJuCrGhBGrQlonGjUk8Ru2
cHqYzT5Sno3M6E6Qtb5zfql7WE0akPjdP1vI7M35SIpvcFhBc6JS7UwmTyHMwJTn
i57BihvtNbLzdKlonC53IfDonb30oLf0FujkArjBnCga+IMy0ZJm5ADbh73CT81p
0POZwfrU31pH4o0pLDMxgarelb4SBiNJBbZm4g8HM4goY6DtRST7389hwBh/Ihjw
gfB6naaFo4AbztdI8EmvHrS2nFHcKFs9NwklvGsE+F/he9xhRL4AYGunX53ToTxk
tJbzNUBSzMTAmaiUH7OJsxM=
'@

# Zertifikatsrotation (2026-09-19): das vorherige Zertifikat (Thumbprint 6CE832BE...) wurde ueber
# eine im Git-Repository committete .pfx-Datei mit LEEREM Passwort kompromittiert (privater
# Schluessel war trivial extrahierbar). Dieses Skript entfernt dessen Thumbprint hier deshalb aktiv
# aus dem Trust-Store, statt das alte Zertifikat einfach weiter als vertrauenswuerdig stehen zu
# lassen - jemand mit dem alten privaten Schluessel koennte sonst weiterhin beliebige Manifeste
# signieren, die auf bereits umgestellten Rechnern trotzdem noch akzeptiert wuerden.
$oldCompromisedThumbprints = @(
    '6CE832BECD40C26E8437D2818F36C512A2BD8A4B'
)

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
        # Alte, kompromittierte Zertifikate immer entfernen (auch im normalen Install-Lauf, nicht
        # nur bei -Uninstall) - siehe "Zertifikatsrotation" oben.
        foreach ($oldThumbprint in $oldCompromisedThumbprints) {
            $oldExisting = $store.Certificates | Where-Object { $_.Thumbprint -eq $oldThumbprint }
            foreach ($oldCert in $oldExisting) {
                $store.Remove($oldCert)
                Write-Host "Altes, kompromittiertes Zertifikat ($oldThumbprint) aus 'CurrentUser\$storeName' entfernt." -ForegroundColor Yellow
            }
        }

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
