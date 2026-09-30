<#
.SYNOPSIS
    Erzeugt ein neues selbstsigniertes ClickOnce-Signaturzertifikat fuer MailDrop und traegt es im Repo ein.

.DESCRIPTION
    Visual Studio signiert die ClickOnce-Manifeste mit dem Zertifikat, dessen Thumbprint in
    MailDrop.vbproj (ManifestCertificateThumbprint) steht, und sucht es im Zertifikatspeicher
    "Eigene Zertifikate" des aktuellen Benutzers (Cert:\CurrentUser\My). Fehlt es dort, bricht der
    Build ab ("Das Manifestsignaturzertifikat wurde nicht im Zertifikatspeicher gefunden").

    Dieses Skript (einmal auf dem Build-Rechner ausfuehren, im Repo-Wurzelordner):
      1. erzeugt ein neues Zertifikat CN=MailDrop (RSA 2048, SHA256, Codesignatur, 30 Jahre gueltig)
         direkt in Cert:\CurrentUser\My - das reicht zum Bauen/Veroeffentlichen,
      2. exportiert es als passwortgeschuetzte .pfx-Sicherung AUSSERHALB des Repos
         (nur zum Wiederherstellen oder fuer einen zweiten Build-Rechner),
      3. traegt den neuen Thumbprint in MailDrop.vbproj ein,
      4. ersetzt das oeffentliche Zertifikat in Publish\Install-Certificate.ps1 und nimmt den
         bisherigen Thumbprint in dessen Entfernen-Liste auf.

    Die .pfx darf NIE ins Repo (siehe .gitignore/CLAUDE.md, Schluessel-Leak 2026-09-19) und auch nicht
    in den Installationsordner auf dem Netzlaufwerk - sie enthaelt den privaten Schluessel.

.PARAMETER PfxPath
    Ziel der .pfx-Sicherung. Standard: Dokumente\MailDrop-Signatur\MailDrop_Signing.pfx

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File tools\signing\New-SigningCertificate.ps1
#>
param(
    [string]$PfxPath = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'MailDrop-Signatur\MailDrop_Signing.pfx')
)

$ErrorActionPreference = 'Stop'
$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..')
$vbproj = Join-Path $repoRoot 'MailDrop.vbproj'
$installScript = Join-Path $repoRoot 'Publish\Install-Certificate.ps1'

$fullPfx = [IO.Path]::GetFullPath($PfxPath)
if ($fullPfx.StartsWith([string]$repoRoot, [StringComparison]::OrdinalIgnoreCase)) {
    throw "Die .pfx-Sicherung darf nicht im Repository liegen: $fullPfx"
}

# Dateien lesen und Zeilenende merken (Windows-Checkout: CRLF, im Repo: LF).
function Read-TextFile($path) {
    $text = [IO.File]::ReadAllText($path)
    $nl = if ($text.Contains("`r`n")) { "`r`n" } else { "`n" }
    return @{ Text = $text; NewLine = $nl }
}
function Write-TextFile($path, $text) {
    [IO.File]::WriteAllText($path, $text, (New-Object Text.UTF8Encoding($false)))
}

$proj = Read-TextFile $vbproj
$m = [regex]::Match($proj.Text, '<ManifestCertificateThumbprint>([0-9A-Fa-f]+)</ManifestCertificateThumbprint>')
if (-not $m.Success) { throw "ManifestCertificateThumbprint nicht in $vbproj gefunden." }
$oldThumbprint = $m.Groups[1].Value.ToUpperInvariant()

Write-Host "Erzeuge neues Zertifikat CN=MailDrop in Cert:\CurrentUser\My ..."
$cert = New-SelfSignedCertificate -Subject 'CN=MailDrop' -Type CodeSigningCert `
    -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 `
    -KeyExportPolicy Exportable -CertStoreLocation 'Cert:\CurrentUser\My' `
    -NotAfter (Get-Date).AddYears(30)
$newThumbprint = $cert.Thumbprint.ToUpperInvariant()
Write-Host "  Thumbprint: $newThumbprint"

$password = Read-Host -AsSecureString 'Passwort fuer die .pfx-Sicherung'
New-Item -ItemType Directory -Force -Path (Split-Path $fullPfx) | Out-Null
Export-PfxCertificate -Cert $cert -FilePath $fullPfx -Password $password | Out-Null
Write-Host "  Sicherung: $fullPfx"

# MailDrop.vbproj: neuer Thumbprint. ManifestKeyFile wird geleert - die .pfx liegt bewusst nicht mehr
# im Projektordner, signiert wird ueber den Zertifikatspeicher.
$projText = $proj.Text.Replace($m.Value, "<ManifestCertificateThumbprint>$newThumbprint</ManifestCertificateThumbprint>")
$projText = [regex]::Replace($projText, '<ManifestKeyFile>[^<]*</ManifestKeyFile>', '<ManifestKeyFile></ManifestKeyFile>')
Write-TextFile $vbproj $projText

# Install-Certificate.ps1: oeffentliches Zertifikat ersetzen, alten Thumbprint zur Entfernen-Liste.
$inst = Read-TextFile $installScript
$nl = $inst.NewLine
$b64 = [Convert]::ToBase64String($cert.Export([Security.Cryptography.X509Certificates.X509ContentType]::Cert))
$lines = for ($i = 0; $i -lt $b64.Length; $i += 64) { $b64.Substring($i, [Math]::Min(64, $b64.Length - $i)) }
$block = '$certBase64 = @''' + $nl + ($lines -join $nl) + $nl + '''@'
$instText = [regex]::Replace($inst.Text, '\$certBase64 = @''\r?\n[\s\S]*?\r?\n''@', { param($x) $block })
if ($instText -notmatch [regex]::Escape($oldThumbprint)) {
    $instText = [regex]::Replace($instText, '(\$oldCompromisedThumbprints = @\(\r?\n)', { param($x) $x.Groups[1].Value + "    '$oldThumbprint'" + $nl })
    # Vorheriger letzter Eintrag braucht jetzt ein Komma nicht - PowerShell-Arrays mit Zeilenumbruch
    # als Trenner sind gueltig, daher keine weitere Anpassung noetig.
}
Write-TextFile $installScript $instText

Write-Host ''
Write-Host 'Fertig. Naechste Schritte:' -ForegroundColor Green
Write-Host '  1. Visual Studio neu laden und bauen (das Zertifikat wird ueber den Thumbprint gefunden).'
Write-Host '  2. Aenderungen an MailDrop.vbproj und Publish\Install-Certificate.ps1 committen (die .pfx NICHT).'
Write-Host '  3. Nach dem naechsten Veroeffentlichen auf jedem Rechner Install-Certificate.ps1 erneut ausfuehren.'
Write-Host '  4. .pfx + Passwort sicher aufbewahren (Passwortmanager), nicht auf dem Netzlaufwerk.'
