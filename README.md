# MailDrop

MailDrop ist ein VSTO-Add-in fuer Microsoft Outlook. Es hilft dabei, E-Mails und optional deren Anhaenge strukturiert in Projektordnern abzulegen.

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

- productive: Stabiler Branch fuer produktive Releases
- development: Integrations-Branch fuer laufende Entwicklung
- feature/*: Kurzlebige Arbeits-Branches, die in development zusammengefuehrt werden
- Freigabe: Reife Staende aus development werden nach productive uebernommen

## Repository-Governance (Best Practice)

- Default-Branch in GitHub: productive
- Pull Request Ziel:
  - Normalfall: feature/* -> development
  - Release-Freigabe: development -> productive
  - Hotfix: hotfix/* -> productive und danach Rueckmerge nach development
- Schutzregeln fuer productive:
  - Keine direkten Pushes
  - Merge nur via Pull Request
  - Mindestens 1 Review erforderlich
  - Alte Reviews bei neuen Commits verwerfen
- Schutzregeln fuer development:
  - Keine direkten Pushes
  - Merge nur via Pull Request
  - Mindestens 1 Review empfohlen

Hinweis: Die eigentliche Default-Branch-Umstellung und Branch-Protection werden in den GitHub-Repository-Settings gesetzt.

## Installation und Setup

### Voraussetzungen

- Windows mit installiertem Outlook Desktop
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

Im Ordner `Publish/` liegen `setup.exe` und das ClickOnce-Manifest `MailDrop.vsto`. VSTO-Add-ins
verlangen zwingend signierte ClickOnce-Manifeste (MSBuild bricht sonst mit
`Cannot build because the ClickOnce manifest signing option is not selected` ab); MailDrop wird
daher mit einem selbstsignierten Zertifikat signiert (kein Zertifikat einer offiziellen
Zertifizierungsstelle). Auf dem Entwicklungsrechner ist dieses Zertifikat bereits vertrauenswuerdig,
auf jedem anderen Rechner schlaegt die Installation daher mit einer Zertifikatswarnung fehl oder
bricht ab.

Abhilfe: Vor der ersten Installation `Publish/Install-Certificate.ps1` ausfuehren. Das Skript vertraut
dem (oeffentlichen) MailDrop-Zertifikat fuer den aktuellen Benutzer, ohne einen privaten Schluessel zu
benoetigen oder zu enthalten, und **ohne Administratorrechte** (Zertifikatsspeicher des aktuellen
Benutzers):

```powershell
powershell -ExecutionPolicy Bypass -File .\Publish\Install-Certificate.ps1
```

Danach `setup.exe` ausfuehren. Das Zertifikat ist bewusst 30 Jahre gueltig (bis 01.07.2056), damit
dieser Trust-Schritt nicht periodisch fuer bereits installierte Benutzer wiederholt werden muss. Nur
falls das Zertifikat jemals neu erzeugt wird (z.B. Kompromittierung des privaten Schluessels), muss
das Skript neu aus `MailDrop.vsto` erzeugt werden (siehe Kommentar im Skript) und alle Benutzer muessen
es erneut ausfuehren.

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

## Lizenz

Aktuell keine explizite Lizenzdatei im Repository hinterlegt.
Falls geplant, bitte LICENSE-Datei ergaenzen.

## Mitwirken

Der verbindliche Entwicklungs- und PR-Workflow ist in CONTRIBUTING.md dokumentiert.
