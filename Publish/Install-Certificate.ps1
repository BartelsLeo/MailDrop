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
MIIC+DCCAeCgAwIBAgIQIruUoIoTVbdGNUQte/MiPjANBgkqhkiG9w0BAQsFADAT
MREwDwYDVQQDDAhNYWlsRHJvcDAgFw0yNjA5MzAxOTM2NTBaGA8yMDU2MDkzMDE5
NDY0OVowEzERMA8GA1UEAwwITWFpbERyb3AwggEiMA0GCSqGSIb3DQEBAQUAA4IB
DwAwggEKAoIBAQCyO6wF4rl+NCgfyDyz556cKMZ1oV84WDRPQuieBEfMZssAxFIr
+7AOA1X5JclF+R1aRBK30gb5loHdM9dVwX/fwNJ4MvJFThixWYPyO78XNa92uDbP
0+jt9LPoFc9blGiEZ+sNpsaS3Vv56bqZ1m5OhhZM37vjZmZCKCgFz3e2DpqjTrpZ
BystzWIvfy0Ph8Hms/qaDpzN/B+6jDxPikHU1XQpuGQyLdRDidJpkzrzAOw7ryOL
DsqPjHf7BV6V0pyO18XXmH8zm5bDUDwmCx6MI3VbXwnHE3epciVEVIge9YyF298I
23ePkrL5gV7Zl7ArsQNwanv0LAr/4AVFUoTJAgMBAAGjRjBEMA4GA1UdDwEB/wQE
AwIHgDATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNVHQ4EFgQU8/beFyK3PGtxrpPS
8EFFWACAwrQwDQYJKoZIhvcNAQELBQADggEBAJggoyxqzoMFf/ukQrWOOCHgPS4k
nBQzo+btiRPfQ7ZIJR57o+lqBk3UtqZnapMOB388bZuQBmPhS9wqcIyMBFcdGbOX
GXoRDwZd+yABx2jumkHNlzoHvo4b0pf4ZfAcW0yp1i9dxO7mYl6tIcRXiSPz2pMU
sAaPesInGPaaAh805hodMYG05B/tfKaan/FoFOLyVzEl8yrISdp+qljCI2Cpf3Tz
xHBTDVsRtoxMF2oUNFsjFBO0ZqGvmCBv8Pgva/I3fncKLTvXxqRx1Q/73Cysu3jg
f2ewfSvv+e/uRzxBseOb7ZMibYG2zVHhv2tdXnjR9HibmuNaObq4ESE79mU=
'@

# Zertifikatsrotation (2026-09-19): das vorherige Zertifikat (Thumbprint 6CE832BE...) wurde ueber
# eine im Git-Repository committete .pfx-Datei mit LEEREM Passwort kompromittiert (privater
# Schluessel war trivial extrahierbar). Dieses Skript entfernt dessen Thumbprint hier deshalb aktiv
# aus dem Trust-Store, statt das alte Zertifikat einfach weiter als vertrauenswuerdig stehen zu
# lassen - jemand mit dem alten privaten Schluessel koennte sonst weiterhin beliebige Manifeste
# signieren, die auf bereits umgestellten Rechnern trotzdem noch akzeptiert wuerden.
$oldCompromisedThumbprints = @(
    '3E84998B1B3EE993EA4F0438B754B6F86F4F7AD1'
    '6CE832BECD40C26E8437D2818F36C512A2BD8A4B'
)

$storeNames = @('Root', 'TrustedPublisher')

function Get-MailDropCertificate {
    $bytes = [Convert]::FromBase64String($certBase64)
    return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(, $bytes)
}

# Internet-Markierung ("Mark of the Web") von allen Dateien dieses Installationsordners entfernen.
# Kommt der Ordner aus einem heruntergeladenen ZIP (GitHub-Release, Browser, Teams, Mail), tragen
# einzelne Dateien die Markierung "aus dem Internet" und andere nicht. ClickOnce bricht dann ab mit
# "Die Bereitstellung und die Anwendung haben keine uebereinstimmenden Sicherheitszonen", weil
# MailDrop.vsto und Application Files\...\MailDrop.dll.manifest verschiedenen Zonen zugeordnet werden.
# Ohne Schreibrechte auf den Ordner (z.B. Netzlaufwerk) schlaegt das fehl - dann nur ein Hinweis.
if (-not $Uninstall) {
    $marked = @(Get-ChildItem -Path $PSScriptRoot -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { Get-Item -LiteralPath $_.FullName -Stream Zone.Identifier -ErrorAction SilentlyContinue })
    if ($marked.Count -gt 0) {
        $failed = 0
        foreach ($file in $marked) {
            try { Unblock-File -LiteralPath $file.FullName -ErrorAction Stop } catch { $failed++ }
        }
        if ($failed -eq 0) {
            Write-Host "Internet-Markierung von $($marked.Count) Datei(en) in diesem Ordner entfernt." -ForegroundColor Green
        } else {
            Write-Warning ("$failed von $($marked.Count) Datei(en) sind als 'aus dem Internet' markiert und konnten nicht " +
                "entsperrt werden (keine Schreibrechte?). Die Installation kann dann mit 'keine uebereinstimmenden " +
                "Sicherheitszonen' scheitern. Abhilfe: das ZIP VOR dem Entpacken entsperren (Rechtsklick > " +
                "Eigenschaften > 'Zulassen') oder den Ordner lokal entpacken und dort installieren.")
        }
        Write-Host ""
    }
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

# Root ist Pflicht (ohne vertrauenswuerdige Stammzertifizierung scheitert die Installation).
# TrustedPublisher ist optional: damit installiert Office ohne Rueckfrage. In Firmenumgebungen
# sperrt eine Richtlinie ("Vertrauenswuerdige Herausgeber nur durch Administratoren verwalten")
# diesen Speicher fuer normale Benutzer ("Zugriff verweigert", erster Feldtest 2026-10-01) - dann
# zeigt der Office-Installer stattdessen eine Rueckfrage, die mit "Installieren" bestaetigt wird.
$rootOk = $true
$publisherOk = $true
foreach ($storeName in $storeNames) {
    try {
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
    catch {
        if ($storeName -eq 'TrustedPublisher') {
            $publisherOk = $false
            Write-Host "'CurrentUser\TrustedPublisher' ist auf diesem Rechner gesperrt (vermutlich Firmenrichtlinie) - nicht schlimm." -ForegroundColor Yellow
        } else {
            $rootOk = $false
            Write-Host "Fehler bei 'CurrentUser\$storeName': $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

Write-Host ""
if ($Uninstall) {
    Write-Host "Fertig. Das MailDrop-Zertifikat wird nicht mehr als vertrauenswuerdig eingestuft." -ForegroundColor Cyan
} elseif (-not $rootOk) {
    Write-Host "Das Zertifikat konnte nicht als vertrauenswuerdig eingetragen werden - die Installation wird" -ForegroundColor Red
    Write-Host "voraussichtlich scheitern. Bitte die IT bitten, das MailDrop-Zertifikat zu verteilen." -ForegroundColor Red
    exit 1
} elseif (-not $publisherOk) {
    Write-Host "Fertig. Jetzt 'setup.exe' aus diesem Ordner ausfuehren. Office fragt dabei einmal nach, ob" -ForegroundColor Cyan
    Write-Host "MailDrop installiert werden soll - mit 'Installieren' bestaetigen." -ForegroundColor Cyan
} else {
    Write-Host "Fertig. Die Installation von MailDrop (setup.exe / MailDrop.vsto) sollte jetzt ohne" -ForegroundColor Cyan
    Write-Host "Zertifikatswarnung funktionieren. Bitte 'setup.exe' aus diesem Ordner erneut ausfuehren." -ForegroundColor Cyan
}
