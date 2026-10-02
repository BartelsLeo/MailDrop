# MailDrop

MailDrop ist ein VSTO-Add-in fuer Microsoft Outlook. Es hilft dabei, E-Mails und optional deren Anhaenge strukturiert in Projektordnern abzulegen.

> **Nur klassisches Outlook fuer Windows.** Das "neue Outlook" (Schalter "Neues Outlook" oben rechts) und Outlook im Browser unterstuetzen keine VSTO-/COM-Add-ins - dort fehlt der MailDrop-Button. Zurueckschalten auf das klassische Outlook bringt ihn zurueck.

## Nutzen

- Schnelleres und konsistentes Ablegen von Outlook-Mails in Projektstrukturen
- Weniger manuelle Fehler durch Platzhalter und Validierung
- Vorschlaege auf Basis historischer Ablagen (SuggestionRecords in SQLite)
- Einheitlicher Ablauf fuer Teams mit wiederkehrenden Projektnamen/-ordnern

## Features

- Ribbon-Button in Outlook zum Oeffnen einer rechten Task Pane
- Bearbeitung nur dann aktiv, wenn genau eine Mail selektiert ist
- Dynamischer Projektstruktur-Baum mit:
  - Neuer Ordner
  - Loeschen (nur leere Ordner)
  - Umbenennen
- Platzhalter-Workflow fuer:
  - Ablageordner
  - msg Dateiname
- Optionale Ablage aller Mail-Anhaenge
- SuggestionEngine mit ONNX-Embeddings fuer Vorschlaege in der Kaskade:
  - ProjektPfad
  - ProjektstrukturPfad
  - Titel
  - Absender (kurz)
  - Ablageordner-Schema
  - msg Dateiname-Schema
  - Anhaenge ablegen (Boolean)
- Sparkle-Hinweise bei automatisch gesetzten Vorschlaegen
- Hilfe-Popup mit:
  - Platzhalter-Referenz
  - Hinweis zum Transfer der SuggestionRecords ueber SQLite-Datei

## Technologie-Stack

- Sprache: Visual Basic .NET
- Framework: .NET Framework 4.7.2
- Add-in-Technologie: VSTO 4.0
- Host: Microsoft Outlook (Desktop)
- Persistenz: SQLite (System.Data.SQLite)
- ML/Embedding: ONNX Runtime + Tokenizer
- Build: Visual Studio 2022 / MSBuild

## Projektstruktur (Kurzueberblick)

- Core/: Session-Logik, Validierung, SuggestionEngine
- Helpers/: DB-Zugriff, TreeView-Helfer, Mail-Helfer
- Services/: EmbeddingService
- UI/: Ribbon, Task Pane, Dialoge, Hilfe-Popup
- Models/: model.onnx, vocab.txt (Runtime-Modellartefakte)

## Branching-Strategie

- main: einziger laufender Branch, direkte Commits fuer Routineaenderungen (frueher development genannt, am 20.09.2026 umbenannt, da er ohnehin der einzige Branch ist)
- feature/*: Kurzlebige Arbeits-Branches fuer isoliert zu review-ende Aenderungen, die in main zusammengefuehrt werden
- Ein frueherer separater released-Branch (Freigabe-Gate development -> released per Pull Request) wurde entfernt, da er einen zusaetzlichen Freigabeschritt ohne ausreichende Release-Kadenz bedeutete; Releases werden jetzt direkt aus main getaggt

## Repository-Governance (Best Practice)

- Default-Branch in GitHub: main
- Pull Request Ziel: feature/* -> main (nur bei Bedarf fuer isoliertes Review, sonst direkter Commit)
- Schutzregeln fuer main:
  - Direkte Pushes fuer Routinearbeit erlaubt
  - Kein Pull Request erforderlich fuer Routineaenderungen

Details zum Vorgehen (inkl. Release-Tagging) siehe CONTRIBUTING.md.

## Installation und Setup

### Voraussetzungen

- Windows mit installiertem Outlook Desktop (klassisches Outlook, nicht das "neue Outlook")
- VSTO Runtime
- Visual Studio 2022
- .NET Framework 4.7.2 Targeting Pack

### Projekt lokal starten

1. Loesung MailDrop.sln in Visual Studio 2022 oeffnen.
2. NuGet-Pakete wiederherstellen (packages.config basiert).
3. Build-Konfiguration Debug | Any CPU waehlen.
4. Projekt starten (F5).
5. Visual Studio startet Outlook als Hostprozess; Add-in wird geladen.

### Installation ueber ClickOnce (Endanwender)

**Schritt-fuer-Schritt-Anleitung und haeufige Fehler: [docs/Installation.md](docs/Installation.md).**
Wichtigster Punkt: Die heruntergeladene ZIP **vor dem Entpacken** freigeben (Rechtsklick >
Eigenschaften > "Zulassen"), sonst bricht die Installation mit "keine uebereinstimmenden
Sicherheitszonen" ab.

Die Distribution erfolgt als entpackbare ZIP-Datei ueber GitHub Releases; dieselbe entpackte ZIP
(`setup.exe`, `MailDrop.vsto`, `Application Files/`) wird zusaetzlich auf ein Netzlaufwerk gelegt.
Wer MailDrop installieren will, kann `setup.exe` entweder aus der lokal entpackten ZIP oder direkt
vom Netzlaufwerk heraus starten.

VSTO-Add-ins verlangen zwingend signierte ClickOnce-Manifeste (MSBuild bricht sonst mit
`Cannot build because the ClickOnce manifest signing option is not selected` ab); MailDrop wird
daher mit einem selbstsignierten Zertifikat signiert (kein Zertifikat einer offiziellen
Zertifizierungsstelle). Auf dem Entwicklungsrechner ist dieses Zertifikat bereits vertrauenswuerdig,
auf jedem anderen Rechner schlaegt die Installation daher mit einer Zertifikatswarnung fehl oder
bricht ab.

Abhilfe: Vor der ersten Installation `Install-Certificate.ps1` ausfuehren (liegt neben `setup.exe`
in der ZIP bzw. auf dem Netzlaufwerk). Das Skript vertraut dem (oeffentlichen) MailDrop-Zertifikat
fuer den aktuellen Benutzer, ohne einen privaten Schluessel zu benoetigen oder zu enthalten, und
**ohne Administratorrechte** (Zertifikatsspeicher des aktuellen Benutzers):

Am einfachsten per Doppelklick auf `Install-Certificate.cmd` (ein Doppelklick auf die `.ps1` oeffnet sie
nur im Editor). Alternativ in PowerShell im selben Ordner:

```powershell
powershell -ExecutionPolicy Bypass -File .\Install-Certificate.ps1
```

Das Skript entfernt ausserdem die Markierung "aus dem Internet" von allen Dateien des Ordners. Ohne
diesen Schritt bricht die Installation aus einem heruntergeladenen ZIP mit "Die Bereitstellung und die
Anwendung haben keine uebereinstimmenden Sicherheitszonen" ab. Liegt der Ordner auf einem Netzlaufwerk
ohne Schreibrechte, das ZIP vor dem Entpacken entsperren (Rechtsklick > Eigenschaften > "Zulassen").

Danach `setup.exe` ausfuehren. Das Zertifikat ist bewusst 30 Jahre gueltig (bis 01.07.2056), damit
dieser Trust-Schritt nicht periodisch fuer bereits installierte Benutzer wiederholt werden muss. Nur
falls das Zertifikat jemals neu erzeugt wird (z.B. Kompromittierung des privaten Schluessels), muss
das Skript neu aus `MailDrop.vsto` erzeugt werden (siehe Kommentar im Skript) und alle Benutzer muessen
es erneut ausfuehren.

#### Updates (kein Auto-Update)

Das automatische Update ist abgeschaltet (`UpdateEnabled=false`, seit 2026-10-02). Grund: Beim
Outlook-Start wurde das Update von `OUTLOOK.EXE` selbst heruntergeladen, und die Defender-Regel
„Office-Anwendungen am Erstellen ausfuehrbarer Inhalte hindern“ blockiert genau das (Meldung
`DeploymentDownloadException` ... `Temp\Deployment\...\MailDrop.dll` „Zugriff verweigert“, siehe
Troubleshooting).

- **Update einspielen:** Outlook schliessen und `setup.exe` der neuen Version (Netzlaufwerk oder neu
  entpackte ZIP) erneut starten. `setup.exe` ist kein Office-Programm und wird von der Regel nicht
  blockiert.
- **Einmalig beim Umstieg:** Installationen einer aelteren Version pruefen beim Outlook-Start noch auf
  Updates und laufen dabei weiter in den Fehler. Einmal `setup.exe` der neuen Version starten (bei
  Problemen vorher deinstallieren) - danach gibt es keine Update-Pruefung mehr.

**Hinweis:** Installation aus einem OneDrive-synchronisierten Ordner (z.B.
`...\OneDrive - <Firma>\...\MailDrop\`) ist keine getestete/unterstuetzte dritte Variante neben
Netzlaufwerk/ZIP. Ob das tatsaechlich (mit-)ursaechlich fuer die unten beschriebenen Fehler ist, ist
nicht belegt (siehe Troubleshooting) - als Vorsichtsmassnahme dennoch vermeiden.

## Verwendung

1. In Outlook eine einzelne Mail auswaehlen.
2. Ribbon-Button MailDrop klicken.
3. Projektverzeichnis waehlen oder anderes... benutzen.
4. Projektstrukturpfad im TreeView auswaehlen/anlegen.
5. Felder pruefen (Titel, Absender kurz, Ablageordner, msg Dateiname).
6. Optional Anhaenge ablegen aktivieren.
7. Mit OK ablegen.

## Platzhalter

Unterstuetzte Platzhalter in Ablageordner und msg Dateiname:

- [Titel]
- [Absender]
- [Absender-Domain]
- [Empfaenger]
- [Empfaenger (kurz)]
- [Betreff]
- [Datum]
- [Datum (formatiert)]
- [Absender (kurz)]

## Datenhaltung und SuggestionRecords

- SQLite-Datei: %APPDATA%/MailDrop/sessions.db
- In dieser Datei liegen die Session- und SuggestionRecords.

### SuggestionRecords auf anderen PC uebernehmen

1. Outlook (und damit MailDrop) auf dem Quell-PC schliessen.
2. Datei %APPDATA%/MailDrop/sessions.db sichern/kopieren.
3. Auf dem Ziel-PC MailDrop/Outlook schliessen.
4. Datei unter exakt demselben Pfad %APPDATA%/MailDrop/sessions.db ablegen und vorhandene Datei ersetzen.
5. Outlook erneut starten.

## Implementierung (Architektur und Ablauf)

### Start und UI

- Einstieg ueber ThisAddIn_Startup
- Ribbon-Button ruft MailAblegen_Click auf
- Task Pane hostet MailDropWpfTaskPane

### Session-Flow

- PrepareSession:
  - Reset
  - Mail-Metadaten lesen
  - Letzte Projektverzeichnisse laden
  - Shared SuggestionEngine beziehen
  - Initiale Feature-Distanzen berechnen
  - ProjektPfad vorschlagen und ggf. uebernehmen
- ProcessSession:
  - Eingaben validieren
  - Zielordner erzeugen
  - Mail als .msg speichern
  - Optional Anhaenge speichern
  - SessionRecord in SQLite speichern

### SuggestionEngine (vereinfacht)

- Historische SessionRecords werden aus SQLite geladen.
- Scoring ueber gewichtete Features, u. a.:
  - Betreff (semantische Aehnlichkeit)
  - Datum
  - Absender/Domain
  - Benutzer
  - ProjektPfad/ProjektstrukturPfad
  - Titel/Ablageordner
- Kaskadierende Vorschlaege entlang des Session-Workflows.

## Validierung und Einschraenkungen

- ProjektPfad muss existieren.
- ProjektstrukturPfad muss unterhalb des ProjektPfads existieren.
- Ungueltige Datei-/Pfadzeichen werden abgefangen.
- Laengenpruefungen fuer Pfade/Dateinamen aktiv.
- Sehr lange Anhangsnamen koennen ueber Umbenennungsdialog behandelt werden.

## Bekannte Hinweise

- Die Gewichtung der SuggestionEngine ist aktuell heuristisch und kann spaeter mit realen Daten feinjustiert werden.
- In mehreren Dateien existieren historische Encoding-Artefakte in Kommentaren/Texten.

## Entwicklungshinweise

- Domainbegriffe und UI-Texte sind bewusst deutsch gehalten.
- Bei Aenderungen an persistierten Daten bitte SQLite-Schema-Auswirkungen mitdenken.
- Bei Aenderungen an Vorschlagslogik Regressionen im Session-Ablagefluss pruefen.

## Verifikation nach Aenderungen

- Build in Visual Studio erfolgreich
- Ribbon-Button oeffnet Task Pane
- Bearbeitung nur bei genau einer selektierten Mail
- Projektstruktur-Baum reagiert korrekt auf Auswahl und Ordneraktionen
- Platzhalter werden bei Fokusverlust aufgeloest
- OK speichert .msg wie erwartet
- Optionales Speichern von Anhaengen funktioniert
- Session wird in SQLite geschrieben

## Troubleshooting

- MailDrop-Button fehlt im Outlook-Ribbon:
  - Pruefen, ob das "neue Outlook" aktiv ist (Schalter oben rechts) - dort laufen keine VSTO-Add-ins; auf das klassische Outlook zurueckschalten.
  - Outlook komplett neu starten.
  - In Outlook unter COM-Add-Ins pruefen, ob MailDrop aktiviert ist.
  - In Visual Studio das Add-in einmal im Debug-Modus starten, damit Registrierung/Load-Verhalten aktualisiert wird.
- Task Pane oeffnet nicht oder bleibt leer:
  - Pruefen, ob genau eine Mail selektiert ist (bei 0 oder >1 ist Bearbeitung deaktiviert).
  - Outlook neu starten und erneut testen.
  - Build auf Debug | Any CPU pruefen und Add-in neu starten.
- Vorschlaege erscheinen nicht:
  - Es werden historische Datensaetze benoetigt (sessions.db darf nicht leer sein).
  - Pruefen, ob Modell-Dateien im Output vorhanden sind: Models/model.onnx und Models/vocab.txt.
  - Outlook einmal neu starten, damit lazy geladene Komponenten frisch initialisiert werden.
- Fehler beim Speichern der Mail/Anhaenge:
  - ProjektPfad und ProjektstrukturPfad auf Existenz pruefen.
  - Dateiname/Pfad auf ungueltige Zeichen oder zu lange Pfade pruefen.
  - Bei langen Anhangsnamen den Umbenennungsdialog verwenden.
- SuggestionRecords wurden auf anderem PC nicht uebernommen:
  - Sicherstellen, dass Outlook auf beiden PCs beim Kopieren geschlossen war.
  - Zielpfad exakt verwenden: %APPDATA%/MailDrop/sessions.db.
  - Vorhandene Datei auf dem Ziel-PC wirklich ersetzen.
- Update schlaegt fehl ("Update nicht moeglich"), obwohl eine Neuinstallation funktioniert, und/oder
  das Add-in wird direkt nach der Installation wegen Zeitueberschreitung deaktiviert, ggf. mit dieser
  Fehlermeldung: `DeploymentDownloadException` / `UnauthorizedAccessException: Der Zugriff auf den
  Pfad "...\AppData\Local\Temp\Deployment\...\MailDrop.dll" wurde verweigert`:
  - **Bestaetigte Ursache beim Update (2026-10-02):** Die Defender-Regel „Office-Anwendungen am
    Erstellen ausfuehrbarer Inhalte hindern“ blockiert den Download, weil das Update beim Outlook-Start
    von `OUTLOOK.EXE` selbst heruntergeladen wird (Windows-Sicherheit → Schutzverlauf: „Aktion
    blockiert“). `setup.exe` ist kein Office-Programm, deshalb klappt eine Neuinstallation. Abhilfe:
    `setup.exe` vom Netzlaufwerk erneut starten oder IT um eine Ausnahme fuer
    `%LOCALAPPDATA%\Temp\Deployment\` bitten.
  - Belegte Ursache (direkt aus dem Stacktrace ablesbar): ClickOnce schreibt heruntergeladene Dateien
    zunaechst in einen zufaellig benannten, rein lokalen Ordner unter
    `%LOCALAPPDATA%\Temp\Deployment\` und oeffnet die Datei danach erneut, um Manifest/Signatur zu
    pruefen. Ein Zugriffsfehler genau in diesem Moment ist ein bekanntes, generisches ClickOnce-
    Problem, meist verursacht durch Virenschutz-/EDR-Software (Echtzeitscan), der neu geschriebene
    DLL/EXE-Dateien kurzzeitig sperrt (der Deployment-Stack wiederholt den Zugriff dabei nicht). Das
    kann unabhaengig vom Installationsort auftreten, da die betroffene Datei lokal liegt.
  - Ein beobachteter Fall stammte von einer Installation aus einem **OneDrive-synchronisierten
    Ordner** (z.B. `...\OneDrive - <Firma>\...\MailDrop\`), keiner der beiden getesteten/
    unterstuetzten Installationsorte (Netzlaufwerk oder woanders entpackte ZIP, siehe oben). Ob
    OneDrive als Quelle tatsaechlich zum Ausloesen des Virenschutz-Zugriffskonflikts beitraegt, ist
    **nicht belegt** - die betroffene Datei liegt lokal und ausserhalb jedes OneDrive-Ordners. Es ist
    daher eher ein Risikofaktor/Grund, bei den unterstuetzten Orten zu bleiben, als eine bewiesene
    Ursache.
  - Dieser Fehler tritt im ClickOnce-/VSTO-Ladevorgang von Outlook auf, bevor der eigene Add-in-Code
    ueberhaupt startet - er ist **nicht** durch eine Codeaenderung in diesem Projekt behebbar.
  - Abhilfe, nach Beleglage geordnet: Virenschutz-/EDR-Ausnahme fuer
    `%LOCALAPPDATA%\Temp\Deployment\` einrichten (die am besten belegte Abhilfe fuer genau diesen
    Fehler), ClickOnce-Cache leeren (`rundll32.exe dfshim.dll CleanOnlineAppCache`), den veralteten
    `DisabledItems`-Registry-Wert fuer MailDrop unter
    `HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems` entfernen, dann neu
    installieren. Zusaetzlich vorsorglich nur vom Netzlaufwerk oder aus einer ZIP ausserhalb von
    OneDrive installieren/aktualisieren, da diese Kombination ungetestet ist.

## Lizenz

Aktuell keine explizite Lizenzdatei im Repository hinterlegt.
Falls geplant, bitte LICENSE-Datei ergaenzen.

## Mitwirken

Der verbindliche Entwicklungs- und PR-Workflow ist in CONTRIBUTING.md dokumentiert.
