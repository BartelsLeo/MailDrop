# MailDrop installieren (Endanwender)

Stand: 2026-10-01, im Feld bestätigt auf einem zweiten Rechner (Installation von OneDrive und vom
Netzlaufwerk `Y:` aus).

## Kurzfassung

1. Das Installationspaket als **ZIP** herunterladen.
2. **Vor dem Entpacken** die ZIP-Datei freigeben: Rechtsklick auf die ZIP → *Eigenschaften* →
   unten bei „Sicherheit“ Haken bei **„Zulassen“** setzen → *OK*.
3. ZIP entpacken (an beliebigen Ort: lokal, OneDrive oder Netzlaufwerk).
4. Im entpackten Ordner **`Install-Certificate.cmd`** doppelklicken. Das Fenster bleibt am Ende offen;
   mit einer beliebigen Taste schließen.
5. **`setup.exe`** starten. Falls Office fragt, ob MailDrop installiert werden soll: **„Installieren“**.
6. Outlook neu starten. Der Button **MailDrop** erscheint im Reiter *Start*.

Pro Benutzer und Rechner einmal nötig.

## Updates

Es gibt **kein automatisches Update** (abgeschaltet, siehe „Häufige Fehler“: Defender blockiert
Updates, die Outlook selbst herunterlädt). Für eine neue Version: Outlook schließen, **`setup.exe`**
der neuen Version erneut starten, Outlook wieder öffnen. Schritt 4 (Zertifikat) ist dafür nicht noch
einmal nötig. Falls `setup.exe` meldet, dass schon eine andere Version installiert ist: MailDrop unter
*Apps & Features* deinstallieren und dann `setup.exe` starten.

## Warum Schritt 2 wichtig ist

Windows markiert heruntergeladene Dateien als „aus dem Internet“. Beim Entpacken einer markierten ZIP
wird diese Markierung je nach Dateityp und Entpackprogramm nur an **einen Teil** der Dateien
weitergegeben. ClickOnce verlangt aber, dass `MailDrop.vsto` und die eigentliche Anwendung
(`Application Files\…`) aus derselben Sicherheitszone stammen. Bei gemischter Markierung bricht die
Installation ab mit:

> Die Bereitstellung und die Anwendung haben keine übereinstimmenden Sicherheitszonen.

Die Markierung wandert beim Kopieren mit – den entpackten Ordner woandershin zu kopieren hilft also
nicht. Wird die ZIP **vor** dem Entpacken freigegeben, sind alle entpackten Dateien unmarkiert und die
Installation funktioniert von jedem Ort aus.

Ausgeschlossene Ursachen (geprüft 2026-10-01): Build-Einstellung (`MailDrop.vsto` enthält keinen
abweichenden `deploymentProvider`) und Speicherort (`Y:`, lokal und OneDrive verhalten sich gleich).

## Wer verteilt: Ordner auf dem Netzlaufwerk befüllen

Damit Benutzer direkt von `Y:\MuP-Water\00_GRL\Software_Setups\MailDrop` installieren können, muss der
Ordner dort **unmarkiert** liegen:

- entweder den `Publish`-Ordner direkt vom Build-Rechner dorthin kopieren (kein Download dazwischen),
- oder die heruntergeladene ZIP zuerst freigeben („Zulassen“) und dann dorthin entpacken.

Prüfen, ob noch markierte Dateien im Ordner liegen (keine Ausgabe = alles in Ordnung):

```powershell
Get-ChildItem -Recurse "Y:\MuP-Water\00_GRL\Software_Setups\MailDrop" |
    Get-Item -Stream Zone.Identifier -ErrorAction SilentlyContinue | Select-Object FileName
```

Nachträglich entsperren (braucht Schreibrechte auf den Ordner):

```powershell
Get-ChildItem -Recurse "Y:\MuP-Water\00_GRL\Software_Setups\MailDrop" | Unblock-File
```

`Install-Certificate.ps1` versucht das ebenfalls automatisch für seinen eigenen Ordner, scheitert aber
ohne Schreibrechte (typisch auf dem Netzlaufwerk) und gibt dann nur einen Hinweis aus.

## Zertifikat (Schritt 4) im Detail

MailDrop ist mit einem selbst erstellten Zertifikat signiert. `Install-Certificate.cmd` startet
`Install-Certificate.ps1`, das dieses Zertifikat für den aktuellen Benutzer einträgt (ohne
Administratorrechte):

- **Vertrauenswürdige Stammzertifizierungsstellen** – Pflicht. Windows fragt dabei evtl. einmal nach;
  mit *Ja* bestätigen.
- **Vertrauenswürdige Herausgeber** – optional. In Firmenumgebungen oft per Richtlinie gesperrt
  („Zugriff verweigert“). Das ist harmlos: Office fragt dann bei `setup.exe` einmal nach, mit
  „Installieren“ bestätigen.

Ein Doppelklick auf die `.ps1` öffnet sie nur im Editor – deshalb die `.cmd`. Alternativ in PowerShell
im Installationsordner:

```powershell
powershell -ExecutionPolicy Bypass -File .\Install-Certificate.ps1
```

Ohne Zutun der einzelnen Benutzer ginge es nur, wenn die IT das Zertifikat zentral per Gruppenrichtlinie
verteilt oder MailDrop mit einem gekauften Code-Signing-Zertifikat signiert wird.

## Häufige Fehler

| Meldung | Ursache | Abhilfe |
|---|---|---|
| „…keine übereinstimmenden Sicherheitszonen“ | Teilweise als „aus dem Internet“ markierte Dateien | ZIP vor dem Entpacken freigeben, neu entpacken, neu installieren (siehe oben) |
| Rote Meldung „Zugriff verweigert“ bei `TrustedPublisher` (ältere Skriptversion) | Firmenrichtlinie sperrt „Vertrauenswürdige Herausgeber“ | Ignorieren, `setup.exe` starten, Rückfrage mit „Installieren“ bestätigen |
| Zertifikats-/Herausgeberfehler bei `setup.exe` | `Install-Certificate` nicht ausgeführt | Schritt 4 ausführen |
| „eine andere Version ist installiert“ | Alte Installation mit früherem Zertifikat | MailDrop unter *Apps & Features* deinstallieren, neu installieren |
| Beim Outlook-Start: `DeploymentDownloadException` … „Zugriff auf den Pfad …\Temp\Deployment\…\MailDrop.dll wurde verweigert“ | Defender-Regel „Office-Anwendungen am Erstellen ausführbarer Inhalte hindern“ blockiert das Update, weil Outlook es selbst herunterlädt (Windows-Sicherheit → Schutzverlauf zeigt „Aktion blockiert“) | Betrifft nur noch ältere Installationen, die selbst nach Updates suchen: einmal `setup.exe` der neuen Version starten (siehe „Updates“) |
| Button fehlt nach Outlook-Neustart | Add-in von Outlook deaktiviert | Siehe README, Abschnitt „Troubleshooting“ |
