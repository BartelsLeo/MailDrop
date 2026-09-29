# MailDrop als Web-Add-in – Umbauskizze

Stand: 2026-09-27 · Status: **Konzeptskizze, nichts davon ist umgesetzt** · Machbarkeitsprüfung mit Recherche: Abschnitt 10

Dieses Dokument skizziert, was nötig wäre, um MailDrop vom heutigen VSTO-Add-in (VB.NET, nur klassisches
Outlook für Windows) auf ein **Office-Web-Add-in** umzustellen, das auch im **neuen Outlook** und in Outlook im
Browser läuft. Es enthält die Grundsatzentscheidungen, die Umbauten pro Komponente, die vorab nötigen
Absprachen mit der IT und die Schritte, die die IT danach durchführt.

> Hinweis: Microsoft ändert Bezeichnungen von Rollen, Berechtigungen und Admin-Oberflächen regelmäßig. Dieses
> Dokument beschreibt, **was** zu tun ist; konkrete Klickpfade, Namen und Grenzwerte sind vor der Umsetzung gegen
> die aktuelle Microsoft-Dokumentation zu prüfen.

---

## 1. Ausgangslage

| | Heute (VSTO) | Web-Add-in |
|---|---|---|
| Läuft in | nur klassischem Outlook für Windows | klassischem Outlook (über WebView2), neuem Outlook, Outlook im Browser, Mac |
| Technik | VB.NET, WPF, .NET Framework 4.7.2 | TypeScript/JavaScript, HTML, Office.js |
| Dateisystem | voller Zugriff (Netzlaufwerk, lokale Ordner) | **kein Zugriff** – läuft abgeschottet wie eine Webseite |
| Verteilung | ClickOnce (`setup.exe`), selbstsigniertes Zertifikat | zentral über das Microsoft-365-Admin-Center |
| Verlauf | SQLite unter `%APPDATA%\MailDrop\sessions.db` | muss an einen erreichbaren Ort (siehe E3) |
| KI-Vorschläge | ONNX Runtime (.NET) | ONNX Runtime Web (im Browser) |

Kernproblem: MailDrops Hauptaufgabe – Dateien in Projektordner schreiben – ist im Web-Add-in nur über die
Microsoft-Schnittstelle **Graph** möglich, also mit **SharePoint/OneDrive als Ablageort**. Deshalb ist das kein
Port, sondern ein Neubau, bei dem die fachliche Logik (Platzhalter, Validierung, Vorschlagskaskade) übernommen
wird.

Ausgangsannahmen aus den Vorgesprächen:

- **Jedes Projekt hat eine eigene SharePoint-Site mit eigener Bibliothek.**
- **Das heutige VSTO-MailDrop ist noch nicht verteilt und wird nicht genutzt.** Es gibt also keinen Übergang,
  keinen Parallelbetrieb, keine bestehenden Ablagen und keinen Altverlauf zu berücksichtigen. Das Web-Add-in
  startet auf der grünen Wiese; das VSTO-Projekt dient nur als fachliche Vorlage (Logik, Oberfläche, Erfahrungen).

---

## 2. Grundsatzentscheidungen

### E1 – Ablageort der Projekte

| Option | Bewertung |
|---|---|
| **SharePoint, eine Site + Bibliothek pro Projekt** (über Graph) | **Empfohlen.** Kein eigener Server nötig, Berechtigungen liegen in SharePoint, passt 1:1 zum heutigen ProjektPfad-Konzept. |
| Netzlaufwerk wie heute | Nur mit **eigenem Server**, der im Auftrag des Add-ins auf die Freigaben schreibt (Anmeldung, Berechtigungsprüfung, Betrieb), oder mit einem **Hilfsprogramm auf jedem PC** (dann entfällt der Hauptvorteil des Web-Add-ins). Nicht empfohlen. |
| Im OneDrive synchronisierte Bibliotheken | Für ein Web-Add-in **nicht nutzbar** – es sieht das lokale Dateisystem nicht. (Nur für das VSTO-MailDrop relevant, siehe Abschnitt 8.) |

### E2 – Zugriff auf die Projekt-Bibliotheken (Berechtigungsmodell)

Grundsatz: Der Nutzer hat über seine SharePoint-Berechtigungen bereits Zugriff auf seine Projekt-Bibliotheken.
Das Add-in ist aber eine **eigene App**, die im Namen des Nutzers auf Graph zugreift – dafür braucht die App
eine **einmalige, tenantweite Zustimmung** (App-Registrierung + Admin-Zustimmung). Diese Zustimmung erweitert
keine Rechte: Die App kann nur, was der angemeldete Nutzer ohnehin darf. Dass eine Bibliothek im OneDrive
verknüpft ist, erteilt der App selbst keine Berechtigung, macht sie aber auffindbar (siehe E4).

| Option | Bewertung |
|---|---|
| **Delegierter Dateizugriff im Namen des Nutzers** (`Files.ReadWrite.All`, delegiert) | **Empfohlen und für den Verknüpfungs-Ansatz (E4) nötig**: Verknüpfte Bibliotheken liegen technisch in fremden Laufwerken, dafür reicht `Files.ReadWrite` (nur eigenes OneDrive) nicht. Effektive Rechte = Nutzerrechte, **keine Freischaltung pro Projekt-Site**, kein Pflegeaufwand bei neuen Projekten. Die Bezeichnung „…All“ ggf. gegenüber der IT erläutern. |
| ~~Freischaltung pro Site (`Sites.Selected`)~~ | Verworfen: Jede Projekt-Site müsste einzeln für die App freigeschaltet werden – unnötiger Pflegeaufwand, da die Zugriffssteuerung ohnehin über die Projektberechtigungen läuft. |

### E3 – Speicherort des Verlaufs („die SQL“)

Der Verlauf (heute `sessions.db`: SessionRecords + ComputedWeights) enthält Betreffzeilen, Absender und
Ablagepfade, also personenbezogene Daten.

| Option | Wandert mit? | Server? | Bewertung |
|---|---|---|---|
| **SQLite-Datei im App-Ordner des Nutzer-OneDrive** | ja | nein | **Empfohlen.** Jede App bekommt dort einen eigenen, privaten Ordner (`/me/drive/special/approot`, Berechtigung `Files.ReadWrite.AppFolder`). Ob das bei Geschäftskonten zuverlässig funktioniert, ist widersprüchlich dokumentiert (Abschnitt 10) – **Rückfallebene ohne Zusatzaufwand:** ein fester Ordner `Apps/MailDrop` im Nutzer-OneDrive, erreichbar über das ohnehin nötige `Files.ReadWrite.All`. Die Datenbank bleibt SQLite (im Browser z. B. über `sql.js`/`wa-sqlite`), das Schema kann weitgehend übernommen werden. Verlauf ist automatisch auf allen Geräten da – der heutige manuelle Export („Vorschlagsdaten exportieren“) entfällt. |
| Browser-Speicher (IndexedDB) | nein | nein | Nur als **lokaler Zwischenspeicher** zusätzlich zu OneDrive; allein zu unsicher (kann gelöscht werden, pro Gerät/Outlook-Variante getrennt). |
| Roaming Settings des Add-ins (im Postfach) | ja | nein | Zu klein (ca. 32 KB). Nur für Einstellungen. |
| Datei/Liste auf SharePoint (gemeinsam fürs Team) | ja | nein | Nur, wenn der Verlauf bewusst **teamweit** geteilt werden soll (neue Kollegen bekommen sofort Vorschläge). Erfordert Regeln, wer welche Betreffzeilen sehen darf. |
| Eigene Datenbank in Azure (+ Azure Functions) | ja | ja (serverlos) | Erst sinnvoll bei Auswertungen über alle Nutzer. Mehr Aufwand, mehr Datenschutzthemen. |

Technische Details zur empfohlenen Variante:

- Beim Öffnen: `sessions.db` aus dem App-Ordner laden (falls vorhanden), lokal in IndexedDB zwischenspeichern.
- Nach jeder Ablage: Datensatz lokal einfügen, Datei zurück in den App-Ordner schreiben.
- **Parallele Nutzung auf zwei Geräten:** Datensätze bekommen eine **GUID**. Vor dem Zurückschreiben wird die
  aktuelle Datei gelesen und per GUID **zusammengeführt** statt überschrieben (Datensätze werden nur angehängt,
  nie geändert – das macht die Zusammenführung einfach). `ComputedWeights` wird lokal neu berechnet, nicht
  zusammengeführt.

### E4 – Projekte finden: Verknüpfungen im persönlichen OneDrive

**Ziel:** Als Projekte werden die Dokumentbibliotheken angeboten, die der Nutzer per
**„Verknüpfung zu ‚Meine Dateien‘ hinzufügen“** in sein OneDrive eingebunden hat; ganz oben die zuletzt
genutzten (aus dem Verlauf).

- Kein Hub, kein Namensschema – jeder sieht genau „seine“ Projekte. Neues Projekt = einmal „Verknüpfung hinzufügen“.
- Gespeichert werden Laufwerk-ID und Element-ID der Bibliothek + Anzeigename statt eines Pfads.

**Achtung – nicht gesichert (Recherche, Abschnitt 10):** Ob Graph diese Verknüpfungen bei **Geschäftskonten**
zuverlässig auflistet, ist widersprüchlich belegt: Verknüpfungen erscheinen als Einträge mit `remoteItem`, es
gibt aber mehrere Berichte, dass sie in `/me/drive/root/children` bei OneDrive for Business **nicht**
zurückgegeben werden. Das ist die **erste Frage, die der Durchstich klären muss.** Lokal nur per
„Synchronisieren“ eingebundene Bibliotheken sind für Graph in keinem Fall sichtbar.

**Rückfallebene (unabhängig von Verknüpfungen), falls die Auflistung nicht klappt:** MailDrop führt eine
**eigene Projektliste** im Verlauf (roamt über OneDrive mit). Ein Projekt wird einmalig hinzugefügt – per
Suche nach Sites über Graph (`/sites?search=`), aus den vom Nutzer **gefolgten Sites** (`/me/followedSites`,
„Stern“ in SharePoint) oder durch Einfügen der Bibliotheks-Adresse – und steht danach dauerhaft in der Liste.
Fachlich gleichwertig, nur ein anderer Weg zum Eintragen.

### E5 – Format der abgelegten Mail

Office.js und Graph liefern eine Mail nur als **.eml** (MIME), **nicht als .msg**. `getAsFileAsync` setzt Requirement Set **Mailbox 1.14** voraus (Abschnitt 10: nicht auf Mobilgeräten; Berichte über Probleme in Outlook 2024 als Kaufversion). Die .eml wird **genauso
abgelegt** wie heute die .msg: gleicher Ablageordner, Dateiname aus dem Dateinamen-Schema, nur mit Endung
`.eml` (Feld „msg Dateiname“ wird zu „Dateiname“). Anhänge können wie heute zusätzlich einzeln abgelegt werden.

- **Inhalt:** Die .eml ist die vollständige Mail inklusive Kopfzeilen, Text und **eingebetteter Anhänge** –
  fachlich gleichwertig zur .msg.
- **Beschaffung:** primär `item.getAsFileAsync()` (Office.js, Mailbox 1.14) direkt aus dem geöffneten Element.
  **Rückfallebene** über Graph (`/me/messages/{id}/$value`, Element-ID per `convertToRestId`), wenn 1.14 im
  jeweiligen Outlook fehlt oder fehlschlägt – dafür zusätzlich `Mail.Read` (delegiert, nur eigenes Postfach).
  Dieselbe Rückfallebene deckt große Anhänge ab (Office.js-Anhangsabruf in neuem Outlook/Web bis ca. 25 MB).
- **Öffnen:** Doppelklick auf eine synchronisierte/heruntergeladene .eml öffnet sie in Outlook (klassisch und
  neu – aktuellen Stand beim neuen Outlook prüfen). In der SharePoint-Weboberfläche wird eine .eml in der Regel
  heruntergeladen statt als Vorschau angezeigt.
- **Upload:** kleine Dateien direkt, große (viele/große Anhänge) per Upload-Session über Graph.
- **Mehrwert gegenüber heute:** Beim Upload können **Metadaten als Spalten der Bibliothek** gesetzt werden
  (Absender, Datum, Betreff, Titel). Dann sind Mails in SharePoint filter- und sortierbar, ohne sie zu öffnen –
  optional, erfordert passende Spalten in den Projektbibliotheken (Site-Vorlage).
- **Sonderfälle prüfen:** verschlüsselte/signierte Mails (S/MIME) und Mails mit Vertraulichkeitsbezeichnung bzw.
  Rechteverwaltung – ob und wie diese als .eml exportierbar und später lesbar sind.

### E6 – Vorschlagsansatz (SuggestionEngine + Embedding) übernehmen

Der bestehende Vorhersageansatz wird **unverändert übernommen**: gewichtete Ähnlichkeit über die Features
(Betreff semantisch, Datum, Absender, Domain, Titel, Ablageordner, ProjektPfad, ProjektstrukturPfad), aus dem
Verlauf gelernte Gewichte (Pearson), Kaskade und geometrischer Neuberechnungs-Trigger.

- **SuggestionEngine (~1.150 Zeilen):** reine Rechenlogik ohne Outlook-/Dateisystem-Bezug → direkte Portierung
  nach TypeScript, per automatischer Tests gegen die Ergebnisse der VB-Version absicherbar. Neuberechnung der
  Gewichte (O(n²)) in einem Web Worker.
- **Betreff-Embedding:** Das heutige Verfahren ist Standard (BERT-WordPiece-Tokenizer aus `vocab.txt`,
  384-dimensionales Modell, Mean Pooling, L2-Normalisierung). **Im Prototyp umgesetzt** (`web/src/embedding/`):
  eigener WordPiece-Tokenizer in TypeScript + ONNX Runtime Web (WASM), getestet in Node und im echten
  Chromium (lokal: 86 MB Download ~3,6 s, Laden ~2,2 s, 3 Embeddings ~0,2 s, plausible Ähnlichkeiten).
- **Gleichstand mit der .NET-Version:** Mit dem deutschfähigen Modell (unten) verwenden VSTO und Web denselben
  Tokenizer und dasselbe Modell; die Gleichheit ist über gemeinsame Prüfvektoren getestet. Für das bisherige
  englische Modell galt das nicht (der Web-Tokenizer entfernt Umlaute, `BertTokenizer` in VSTO vermutlich
  nicht) – dank fehlendem Altverlauf war das kein Problem.
- **Modellwahl – Entscheidung 2026-09-27: ein deutschsprachiges Modell genügt.** Das heutige Modell ist rein
  englisch (`vocab.txt` = englisches uncased-BERT-Vokabular, 30.522 Einträge); deutsche Betreffzeilen werden in
  viele Wortstücke zerlegt, die semantische Ähnlichkeit leidet.
  - Rein deutsche Satzmodelle gibt es praktisch nur in großen Varianten (BERT-large-Größe, mehrere hundert MB)
    – für den Aufgabenbereich ungeeignet.
  - **Geplant:** `paraphrase-multilingual-MiniLM-L12-v2` (118 Mio. Parameter, 384 Dimensionen – gleiche
    Schnittstelle wie heute). In einem deutschen Vergleich deutlich besser als GBERT-large (Korrelation 0,84 vs.
    0,67). Das Modell ist nur deshalb groß (~470 MB), weil ~96 Mio. Parameter auf das Vokabular für 50+ Sprachen
    entfallen. Da Deutsch genügt, wird das **Vokabular auf deutsch (+ englisch) relevante Tokens gekürzt** und
    das Modell auf 8 Bit quantisiert → erwartet ca. 20–40 MB. Die Gewichte der behaltenen Tokens bleiben
    unverändert, die Qualität für deutsche Texte damit praktisch gleich (per Vergleich mit dem Originalmodell
    auf echten Betreffzeilen zu bestätigen).
  - **Umsetzung (2026-09-27, für VSTO und Web gemeinsam):** Beide Add-ins nutzen dasselbe gekürzte Modell aus
    `Models/`, damit die Vorschläge auf beiden Seiten gleich funktionieren.
    - `tools/model/build_german_model.py` lädt das Originalmodell (ONNX-Fassung aus dem Hugging-Face-Repo),
      zählt die Tokens auf einem Korpus aus 40.000 deutschen + 5.000 englischen Wikipedia-Artikelanfängen,
      behält die häufigen Tokens (höchstens 50.000) plus alle Einzelzeichen der lateinischen Schriften,
      schneidet die Token-Tabelle zu und quantisiert auf int8. Ausgabe unter den **bisherigen Dateinamen**
      `Models/model.onnx` und `Models/vocab.txt` (jetzt `Token<TAB>Score` je Zeile), dazu
      `testvectors.json` (Prüfvektoren) und `MODEL_INFO.md` (Revision, Prüfsummen, Messwerte).
    - Der Tokenizer (SentencePiece/Unigram) ist dreimal gleich umgesetzt: `web/src/embedding/unigram.ts`,
      `Services/UnigramTokenizer.vb`, Referenz im Build-Skript. Kein Transformers.js – eine eigene, kleine
      Umsetzung lässt sich in VB genauso schreiben und über dieselben Prüfvektoren absichern.
    - Das Skript bricht ab, wenn der eigene Tokenizer auf weniger als 99 % der Korpustexte vom gekürzten
      Hugging-Face-Tokenizer abweicht oder die Embeddings im Mittel unter Kosinus 0,97 zum Originalmodell
      fallen. Danach prüfen Web-Tests (Vitest) und `tools/tokenizer-check` (VB) die Prüfvektoren.
    - Läuft in GitHub Actions (`.github/workflows/build-german-model.yml`, bei Änderungen unter
      `tools/model/` oder per Hand) und checkt das Ergebnis ein – Hugging Face ist aus der Claude-Umgebung
      gesperrt. Das Modell bleibt eingecheckt (erwartet 30–45 MB, unter der 100-MB-Grenze), damit VSTO-Build
      und ClickOnce ohne Download-Schritt auskommen.
    - `EmbeddingService.vb` und `embedder.ts` erkennen das Format von `vocab.txt`: bis das neue Modell
      eingecheckt ist, läuft unverändert das englische Modell weiter.
  - **Azure Static Web Apps:** Das kleinere Modell entschärft zugleich die Speichergrenzen (Free-Tarif 250 MB
    gesamt; gemeldete Grenze von 100 MB pro Datei – das heutige 86-MB-Modell läge knapp darunter).
- **Offener Punkt Laufzeit:** Ladezeit beim ersten Start und Speicherverbrauch des (gekürzten, int8-)Modells im
  Aufgabenbereich des (neuen) Outlook im Pilot messen.

---

## 3. Technischer Aufbau (Zielbild)

- **Frontend:** TypeScript, React + Fluent UI (Office-Optik), gebaut als statische Seite.
- **Hosting:** Azure Static Web Apps (statisch, kein eigener Server), automatische Veröffentlichung per GitHub Actions.
- **Manifest:** Add-in-Beschreibung (XML-Manifest oder das neuere einheitliche JSON-Manifest – aktuellen
  Unterstützungsstand für Outlook prüfen). Button im Lesebereich und im geöffneten Mail-Fenster
  (Befehlsoberfläche „Nachricht lesen“).
- **Anmeldung:** MSAL.js mit **Nested App Authentication** (von Microsoft für Office-Add-ins empfohlen, Anmeldung
  über das Outlook-Konto ohne eigenes Anmeldefenster, **kein Client-Geheimnis**).
- **Daten:** Microsoft Graph für SharePoint (Projekte, Ordner, Upload) und OneDrive-App-Ordner (Verlauf).
- **Rechenintensives** (Embedding, Gewichtsneuberechnung O(n²)) in einem **Web Worker**, damit die Oberfläche
  nicht einfriert.

---

## 4. Umbauten pro Komponente

| Heute | Web-Add-in | Art |
|---|---|---|
| `ThisAddIn.vb` – Lebenszyklus, Explorer-Events, Task Pane, Preloads | Task Pane mit **Anheften** (`SupportsPinning`) + Ereignis `ItemChanged` statt `SelectionChange`. Startzeit-Probleme, ClickOnce-Diagnose, COM-Freigabe-Fallen entfallen komplett. | Neubau, stark vereinfacht |
| Ribbon (`MailDropRibbon.*`, Explorer + Inspector) | Button im Manifest; erscheint automatisch im Lesebereich **und** im geöffneten Mail-Fenster. Die „feste Bindung“ an ein Mail-Fenster entfällt. | entfällt/ersetzt |
| Toggle-Button, Ausblenden nach Ablage | Task Pane öffnen/schließen über Office; Schließen per Code nur eingeschränkt möglich (aktuellen Stand prüfen). Überlagerung der Leseansicht ist auch hier nicht möglich (Pane ist angedockt). | anpassen |
| `MailDropWpfTaskPane.xaml(.vb)` | React-Komponenten mit Fluent UI; gleiche Felder und Reihenfolge. | Neubau |
| TreeView ProjektstrukturPfad (`DirectoryTreeHelper`) | Baum über Graph, **lädt nur die aufgeklappte Ebene** (löst nebenbei das heutige Performance-Problem des vollständigen Einlesens). Neuer Ordner/Löschen/Umbenennen über Graph. Synthetischer Wurzelknoten bleibt. | Neubau |
| ProjektPfad-Liste + „anderes…“ | Letzte Projekte aus dem Verlauf + Projektsuche (E4). Gespeichert werden Site-ID/Bibliothek-ID + Anzeigename statt Pfad. | Neubau |
| `Session.vb` – Zustand, Kaskade, Platzhalter (`ReplacePlaceholders`) | TypeScript-Zustand; Platzhalter-Logik (Literal/Connector/Placeholder, Skip-Empty-Join) **1:1 portieren**, mit Unit-Tests aus den dokumentierten Beispielen. | Portierung |
| `InputChecker.vb` | Portieren, aber **SharePoint-Regeln** statt Windows-Regeln: verbotene Zeichen, führende/abschließende Leerzeichen, reservierte Namen, maximale Pfadlänge (URL-dekodiert ca. 400 Zeichen), Dateigröße. `IsInsideBaseFolder` sinngemäß über Bibliothekspfade. | Portierung + Anpassung |
| `MailUtils.vb` – Metadaten, .msg, Anhänge | Office.js: Betreff, Absender, Empfänger, Datum aus `item`; Mail als .eml über `getAsFileAsync`; Anhänge über `getAttachmentContentAsync`; Upload über Graph (kleine Dateien direkt, große per Upload-Session). | Neubau |
| `AttachmentRenameDialog` | Dialog innerhalb der Pane. | Neubau |
| `DatabaseUtils.vb` / SQLite / Migrationen | SQLite im Browser + Sync mit OneDrive-App-Ordner (E3). Schema übernehmen, **GUID-Spalte ergänzen**; `PRAGMA user_version`-Migrationen weiterverwenden. | Portierung + Sync neu |
| `SuggestionEngine.vb` | Portieren (Features, Gewichte, Kaskade, geometrischer Neuberechnungs-Trigger, Pearson-Gewichte); Neuberechnung im Web Worker. Ergebnisse gegen die .NET-Version mit denselben Daten abgleichen. | Portierung |
| `EmbeddingService.vb` + `Models/` | ONNX Runtime Web + Tokenizer in TS; Modell auf dem Webspace, Browser-Cache. | Portierung |
| `NotificationToast` | Meldung innerhalb der Pane oder Office-Benachrichtigungsleiste am Element; „Öffnen“ öffnet den Ablageordner in SharePoint (Browser) statt im Explorer. | Neubau |
| `InfoPopup` | Panel/Dialog in der Pane; „Vorschlagsdaten exportieren“ entfällt (Verlauf liegt in OneDrive). | Neubau, vereinfacht |
| `Logger` / `error.log` | Entscheidung: Browser-Konsole + Fehleranzeige, oder Application Insights in Azure (dann **ohne** Mail-Inhalte loggen, Datenschutz). | Neubau |
| ClickOnce, `Install-Certificate.ps1`, Zertifikat, `SQLite.Interop.dll`-Workaround | **Entfallen** vollständig. | entfällt |

---

## 5. Vorgehen in Phasen

1. **Absprachen mit der IT** (Abschnitt 6) und Entscheidungen E1–E6 festhalten.
2. **Technischer Durchstich (Prototyp):** **Angelegt in `web/`** (Branch `feature/web-prototype`, Anleitung
   `web/README.md`) als Prüfoberfläche mit neun Prüfungen: Umgebung/Requirement Sets, Anmeldung (NAA),
   aktuelle Mail/freigegebenes Postfach, .eml-Export (Office.js + Graph), Anhänge, Projektbibliotheken
   (Verknüpfungen + gefolgte Sites), Verlaufsspeicher (App-Ordner + Rückfallebene), Test-Upload, Modell.
   Beantwortet gezielt die offenen Punkte aus Abschnitt 10; Ergebnis als kopierbarer Bericht.
3. **Kernfunktionen:** komplette Oberfläche, Platzhalter, Validierung, Ordneraktionen, Verlauf in OneDrive.
4. **Vorschläge:** SuggestionEngine + Embedding portieren, Abgleich mit der .NET-Version.
5. **Pilot** mit kleiner Gruppe.
6. **Rollout** an alle.

---

## 6. Absprachen mit der IT (vorab zu klären)

**Grundsätzliches**

1. Ist **SharePoint als Ablageort** für Projekt-Mails gewollt und verbindlich? (Ohne das lohnt der Umbau nicht.)
2. Liegen **alle** Projekte (auch laufende) bereits auf SharePoint, oder gibt es noch Projektordner auf dem Netzlaufwerk, in die abgelegt werden soll?
3. Wann ist mit einem **verpflichtenden Umstieg auf das neue Outlook** zu rechnen? (Bestimmt den Zeitdruck.)

**Projekt-Sites**

4. Ist es in Ordnung, dass Nutzer ihre Projektbibliotheken per **„Verknüpfung zu ‚Meine Dateien‘ hinzufügen“** in ihr OneDrive einbinden (ist diese Funktion im Tenant aktiv)?
5. Wie werden Projekt-Sites **angelegt** (Vorlage)? Relevant für optionale Metadaten-Spalten (E5).
6. Gibt es **Aufbewahrungs-/Compliance-Richtlinien** für abgelegte Mails?

**Berechtigungen und Anmeldung**

7. Ist **delegierter Dateizugriff „im Namen des Nutzers“** (`Files.ReadWrite.All`, delegiert) zulässig? (Effektiv nur Nutzerrechte; nötig, um auf verknüpfte Bibliotheken zuzugreifen.)
8. Sind **`Files.ReadWrite.AppFolder`** (Verlauf) und **`Mail.Read`** (delegiert; Rückfallebene für .eml-Export und große Anhänge, nur eigenes Postfach) zulässig?
9. Gibt es **Richtlinien für bedingten Zugriff** (verwaltete Geräte, MFA, Standorte), die das Add-in betreffen?

**Azure und Verteilung**

10. Darf eine **Azure Static Web App** im Tenant betrieben werden? Welche Subscription/Ressourcengruppe, wer ist Besitzer?
11. Eigene Adresse (z. B. `maildrop.firma.de`) gewünscht?
12. Wer verteilt das Add-in im **Admin-Center (Integrierte Apps)**, an welche Gruppen?
13. Welche Outlook-Varianten sind im Einsatz (klassisch/neu/Web/Mobil)? Klassisches Outlook als **Microsoft-365-Abo (aktueller Kanal)** oder als **Kaufversion (2021/2024/LTSC)**? Davon hängt ab, ob Mailbox 1.14 verfügbar ist. Wird aus **freigegebenen Postfächern/Stellvertretungen** abgelegt? (Unterstützt; dann zusätzlich `Mail.Read.Shared`.)
14. Müssen Adressen im **Proxy/Firewall** freigegeben werden?

**Datenschutz**

15. Datenschutz-Prüfung / Eintrag im Verzeichnis der Verarbeitungstätigkeiten: Verlauf mit Betreff, Absender, Pfaden im OneDrive des Nutzers; KI läuft lokal; keine Daten an Dritte.
16. Verlauf **pro Person** (OneDrive) oder **teamweit** (SharePoint)?

---

## 7. Schritte der IT (nach den Absprachen)

**Azure**

- [ ] Ressourcengruppe bereitstellen, **Azure Static Web App** anlegen (kleinster Tarif genügt).
- [ ] Veröffentlichung einrichten: Entwickler als Mitwirkender auf der Ressourcengruppe **oder** Bereitstellungstoken für GitHub Actions herausgeben.
- [ ] Optional: eigene Domain verbinden (Zertifikat stellt Azure automatisch).

**Entra ID (App-Registrierung)**

- [ ] App-Registrierung „MailDrop“ anlegen (nur eigener Tenant).
- [ ] Plattform „Single-Page-Anwendung“ mit den Umleitungs-URIs für Nested App Authentication (Format laut aktueller Microsoft-Doku, u. a. `brk-multihub://<add-in-domain>`) und der Add-in-Adresse.
- [ ] Delegierte Graph-Berechtigungen eintragen: `User.Read`, `Files.ReadWrite.All`, `Files.ReadWrite.AppFolder`, `Mail.Read`, bei Ablage aus freigegebenen Postfächern zusätzlich `Mail.Read.Shared`; `Sites.Read.All` für die Rückfallebene „gefolgte Sites“ bei der Projektsuche (E4).
- [ ] Für den Prototyp (`web/`): Umleitungs-URIs `brk-multihub://localhost:3000` und `https://localhost:3000/taskpane.html`; Hochladen eigener Add-ins für den Testnutzer erlauben oder Manifest `web/manifest.xml` an ihn verteilen.
- [ ] **Administrator-Zustimmung** erteilen.
- [ ] Richtlinien für bedingten Zugriff prüfen/anpassen.
- [ ] Kein Client-Geheimnis/Zertifikat nötig – entsprechend kein Ablaufdatum zu überwachen.

**SharePoint**

- [ ] Sicherstellen, dass „Verknüpfung zu ‚Meine Dateien‘ hinzufügen“ für die Projektbibliotheken verfügbar ist (E4); Nutzer kurz anleiten.
- [ ] Berechtigungen pro Projekt-Site wie gewohnt pflegen (bleibt die eigentliche Zugriffssteuerung).
- [ ] Ggf. Aufbewahrungsrichtlinien für die Bibliotheken festlegen.
- [ ] Optional: Spalten für Mail-Metadaten (Absender, Datum, Betreff) in die Vorlage der Projektbibliotheken aufnehmen (E5).

**Verteilung und Betrieb**

- [ ] Manifest im **Microsoft-365-Admin-Center → Integrierte Apps** hochladen, zunächst an eine **Pilotgruppe** zuweisen, später an alle.
- [ ] Sicherstellen, dass Add-ins in Outlook nicht per Richtlinie gesperrt sind.
- [ ] Proxy/Firewall: Adresse der Static Web App freigeben.
- [ ] Outlook-Versionen/WebView2 auf den PCs prüfen.
- [ ] Datenschutz-Dokumentation abschließen.
- [ ] Laufend: neue Nutzer der Gruppe hinzufügen, neue Projekt-Sites anlegen (und ggf. freischalten). Updates des Add-ins verteilen sich automatisch mit jeder Veröffentlichung.

---

## 7a. Aufwand (Umsetzung vollständig durch Claude Code)

Der Code wird vollständig von Claude Code geschrieben. Der Aufwand verschiebt sich damit vom Programmieren auf
**Prüfen und Testen in der echten Microsoft-365-Umgebung**, das Claude nicht selbst kann (kein Zugriff auf
euren Tenant, kein Outlook).

| Was | Wer | Einschätzung |
|---|---|---|
| Code schreiben (Oberfläche, Graph-Zugriff, Verlauf, Portierung Session/Platzhalter/Prüfungen/SuggestionEngine) | Claude | kein Engpass |
| Automatische Tests der portierten Logik (Platzhalter-Beispiele, Validierung, Vorschlagsberechnung, Embedding-Abgleich) | Claude, im eigenen Container | kein Engpass – deckt die fachlich kritischen Teile ab |
| IT-Einrichtung (Azure, App-Registrierung, Zustimmung, Verteilung) | IT | wenig Arbeit, aber **Wochen Vorlauf** möglich |
| Testen im echten Outlook/SharePoint, Fehler zurückmelden | Nutzer | **eigentlicher Engpass**: realistisch ca. 3–5 Testrunden für den Durchstich, 10–20 Runden bis zur fertigen Version |
| Pilot und Nachbesserungen | Nutzer + Claude | einige Runden |

Grobe Kalenderdauer bei zügigen Testrunden: **Durchstich in wenigen Tagen** nach Abschluss der IT-Einrichtung,
**fertige Version in einigen Wochen** – bestimmt durch die Zahl der Testrunden und den IT-Vorlauf, nicht durch
das Schreiben des Codes. Unsicherheit bleibt v. a. bei Anmeldung/Berechtigungen (hängt an IT-Einstellungen) und
beim Verhalten des Modells im Outlook-Aufgabenbereich (Ladezeit/Speicher).

---

## 8. Alternative: VSTO-MailDrop fertigstellen statt Web-Add-in

Da das VSTO-MailDrop noch nicht verteilt ist, steht die Grundsatzfrage jetzt: **VSTO fertigstellen oder direkt
das Web-Add-in bauen?**

| | VSTO fertigstellen | Web-Add-in bauen |
|---|---|---|
| Aufwand bis zum Einsatz | gering (existiert, muss getestet und verteilt werden) | hoch (Neubau) |
| Läuft im neuen Outlook | nein | ja |
| Lebensdauer | begrenzt durch Supportende des klassischen Outlook | zukunftssicher |
| Verteilung | ClickOnce, selbstsigniertes Zertifikat, bekannte Probleme (Deaktivierung wegen langsamen Starts, `Temp\Deployment`-Fehler) | zentral über Admin-Center, keine Installation |
| SharePoint als Ablageort | nur über im OneDrive synchronisierte Bibliotheken (lokale Ordner, jeder Nutzer muss synchronisieren) | direkt über Graph |
| IT-Aufwand | Zertifikat/Verteilung | Azure, App-Registrierung, Admin-Center (Abschnitt 7) |

Mit VSTO wäre SharePoint nur über synchronisierte Bibliotheken nutzbar: Jeder Nutzer synchronisiert die
Projektbibliothek über OneDrive; sie erscheint als lokaler Ordner und kann über „anderes…“ als ProjektPfad
gewählt werden, OneDrive lädt abgelegte Dateien hoch. Erweiterungen dafür wären: synchronisierte Bibliotheken
automatisch in der ProjektPfad-Liste anbieten, Pfadlängenprüfung an SharePoint-Grenzen anpassen.

Einschätzung: Wenn der Umstieg auf das neue Outlook in absehbarer Zeit kommt oder die Projekte ohnehin auf
SharePoint liegen, spricht viel dafür, **das VSTO-MailDrop gar nicht erst zu verteilen** und die Energie direkt
in das Web-Add-in zu stecken – sonst wird zweimal ausgerollt und zweimal geschult.

---

## 9. Risiken und offene Punkte

- **.eml statt .msg** (E5) – fachlich gleichwertig; Sonderfälle verschlüsselter/gekennzeichneter Mails prüfen.
- **Berechtigungsmodell** (E2) – lehnt die IT `Files.ReadWrite.All` (delegiert) ab, funktioniert der Verknüpfungs-Ansatz nicht; Ausweichen wäre Freischaltung pro Site mit Automatisierung.
- Die in Abschnitt 10 recherchierten Fallstricke (Verknüpfungen, Mailbox 1.14, App-Ordner, Speicher).
- **Erster Start** mit ~90 MB Modell – ggf. kleineres Modell nötig.
- **Vorschlagsqualität** nach Portierung – nur durch Abgleich mit der .NET-Version auf denselben Daten nachweisbar.
- **Graph-Drosselung** bei sehr großen Bibliotheken/Ordnerbäumen – durch ebenenweises Laden entschärft.
- **Office.js-Funktionsumfang** (Anheften, Schließen der Pane per Code, Mehrfachauswahl) hängt von den unterstützten Requirement Sets der eingesetzten Outlook-Versionen ab – vor Phase 2 prüfen.

---

## 10. Machbarkeitsprüfung und Fallstricke (Recherche 2026-09-27)

**Gesamturteil: umsetzbar, kein K.-o.-Kriterium gefunden.** Drei Punkte sind aber nicht durch Recherche zu
klären und müssen im Durchstich als Erstes praktisch geprüft werden; für jeden gibt es eine Rückfallebene.

### Bestätigt

| Annahme | Ergebnis |
|---|---|
| Neues Outlook unterstützt keine COM-/VSTO-Add-ins | Bestätigt. |
| Zeitdruck durch das neue Outlook | Microsoft unterstützt das klassische Outlook bis mindestens 2029 (M365 und LTSC). Für Unternehmen ist aber eine **Opt-out-Phase ab März 2027** angekündigt, der endgültige Wechsel **nicht vor März 2028**, mit mindestens 12 Monaten Vorankündigung. → Ein VSTO-MailDrop hätte im M365-Umfeld realistisch nur ca. 1,5–3 Jahre Lebensdauer. |
| Anmeldung per Nested App Authentication (NAA) | Bestätigt und inzwischen **Pflicht**: Die alten Exchange-Token für Add-ins sind in allen Tenants abgeschaltet. Umleitungs-URI `brk-multihub://<domain>` (nur Origin, kein Pfad), SPA-Plattform, MSAL.js mit `createNestablePublicClientApplication`. Verfügbarkeit per `isSetSupported("NestedAppAuth", "1.1")` prüfen und Rückfallebene (Dialog-Anmeldung) vorsehen. |
| Angeheftete Pane + `ItemChanged` | Unterstützt in neuem Outlook, Web, klassischem Outlook (M365) und Mac. **Ohne Anheften schließt sich die Pane beim Wechsel der Mail** – Nutzer müssen einmal auf die Stecknadel klicken; der `ItemChanged`-Handler muss `item === null` behandeln (z. B. Mehrfachauswahl). |
| .eml wird im neuen Outlook geöffnet | Seit März 2024 unterstützt (Doppelklick, wenn das neue Outlook Standard-App für .eml ist; sonst „Öffnen mit“). |
| SharePoint-Grenzen | Gesamtpfad (dekodiert) max. **400 Zeichen**, einzelner Name max. 255; verboten `* " : < > ? / \ |`, führende/abschließende Leerzeichen, reservierte Namen (`CON`, `PRN`, `AUX`, `NUL`, `COM0–9`, `LPT0–9`, `_vti_`, `desktop.ini`, `~$…`, `.lock`). → Neue Regeln für `InputChecker`. |
| Upload | Direkt bis 250 MB, darüber Upload-Session (Teile < 60 MiB, Vielfaches von 320 KiB) – für Mails unkritisch. |
| Add-in nicht auf SharePoint hostbar | Bestätigt: Outlook-Add-ins werden von SharePoint-App-Katalogen nicht unterstützt; Webspace (Azure Static Web Apps) nötig. |

### Offen – im Durchstich zu prüfen

1. **Auflistung der OneDrive-Verknüpfungen (E4) – größtes Risiko.** Verknüpfungen werden als Einträge mit
   `remoteItem` beschrieben, es gibt aber mehrere Berichte, dass sie bei OneDrive for Business **nicht** in
   `/me/drive/root/children` erscheinen. → Rückfallebene: eigene Projektliste in MailDrop (Site-Suche,
   gefolgte Sites, Adresse einfügen), siehe E4.
2. **`getAsFileAsync` / Mailbox 1.14 (E5).** Nicht auf Mobilgeräten; offener Fehlerbericht, dass es in
   **Outlook 2024 (Kaufversion)** trotz dokumentierter Unterstützung „nicht unterstützt“ meldet; im klassischen
   Outlook weichen die erzeugten Kopfzeilen teils ab. → Rückfallebene Graph-MIME-Export mit `Mail.Read`.
   IT-Frage 13 (Abo vs. Kaufversion) klärt, wie relevant das ist.
3. **Modell im Aufgabenbereich (E6).** Es gibt Berichte über hohen Speicherverbrauch/Speicherlecks von WebView2
   in Office-Add-ins; ONNX Runtime Web läuft zuverlässig single-threaded über WASM. → Modell erst bei Bedarf in
   einem Web Worker laden, Ladezeit/Speicher messen, ggf. quantisierte Variante (~¼ Größe).

### Korrigiert gegenüber der ersten Skizze

- **App-Ordner (E3):** Die aktuelle Graph-Doku beschreibt den App-Ordner für OneDrive privat **und** geschäftlich
  mit `Files.ReadWrite.AppFolder`; ältere Microsoft-Antworten sagen „nur private Konten“. → Nicht mehr als
  gesichert behandelt; Rückfallebene fester Ordner `Apps/MailDrop` über `Files.ReadWrite.All`.
- **`Mail.Read` zusätzlich beantragen:** Als Rückfallebene für den .eml-Export und für große Anhänge
  (Office.js-Anhangsabruf in neuem Outlook/Web ca. 25 MB). Delegiert, nur eigenes Postfach.

### Weitere Fallstricke für die Umsetzung

- **Max. 3 gleichzeitige asynchrone Office.js-Aufrufe** in neuem Outlook/Web → rein interne Programmierregel: Beim **einen** Klick auf OK holt der Code die Anhänge in einer Warteschlange nacheinander (bzw. max. 3 parallel) ab. Für den Nutzer ändert sich nichts – weiterhin ein Ablagevorgang pro Mail inkl. aller Anhänge.
- **Roaming Settings max. 32 KB** → nur für kleine Einstellungen, nicht für den Verlauf (wie geplant).
- **Browser-Speicher (IndexedDB) kann gelöscht werden** → nur Zwischenspeicher, maßgeblich ist die Datei in OneDrive (wie geplant).
- **Freigegebene Postfächer/Stellvertretung:** Wird unterstützt – der Nutzer hat ja Zugriff. Es ist nur ein anderer technischer Weg: Office.js erkennt das freigegebene Postfach (`getSharedPropertiesAsync`), die Graph-Rückfallebene greift über `/users/{postfach}/messages` statt `/me` zu und braucht dafür zusätzlich `Mail.Read.Shared` (delegiert, nur Postfächer, auf die der Nutzer ohnehin Zugriff hat). Im Durchstich mit einem echten freigegebenen Postfach testen. Der Vorschlags-Verlauf bleibt dabei der persönliche des Nutzers.
- **Geschützte Mails** (S/MIME, Vertraulichkeitsbezeichnungen/Rechteverwaltung): Exportierbarkeit als .eml und spätere Lesbarkeit im Durchstich prüfen.
- **Zielplattformen:** Desktop (neues und klassisches Outlook) **und Outlook im Browser** – der Browser ist technisch praktisch dasselbe wie das neue Outlook (Mailbox 1.14, Anmeldung, Anheften, Modell im Browser). Einziger Browser-spezifischer Prüfpunkt: gemeldete Probleme der Anmeldung (NAA), wenn im Browser Drittanbieter-Cookies blockiert sind → im Durchstich testen, Rückfallebene Dialog-Anmeldung. **Mobil ist kein Ziel** (kein Mailbox 1.14).

### Quellen

- Mailbox 1.14 / `getAsFileAsync`: [Requirement Set 1.14](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/outlook/requirement-set-1.14/outlook-requirement-set-1.14?view=common-js-preview) · [Issue #5582 (Outlook 2024)](https://github.com/OfficeDev/office-js/issues/5582) · [Issue #5635 (Kopfzeilen klassisch)](https://github.com/OfficeDev/office-js/issues/5635) · [Q&A: kein 1.14 auf Mobil](https://learn.microsoft.com/en-ca/answers/questions/5624660/what-are-the-api-options-for-retrieving-full-eml-i)
- NAA: [Enable NAA](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in) · [NAA-FAQ / Legacy-Token](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/faq-nested-app-auth-outlook-legacy-tokens)
- Verknüpfungen: [remoteItem](https://learn.microsoft.com/en-us/graph/api/resources/remoteitem?view=graph-rest-1.0) · [Q&A: Shortcuts fehlen in children](https://learn.microsoft.com/en-us/answers/questions/761174/microsoft-graph-api-how-to-get-shortcuts-when-gett) · [Tech Community: OneDrive Shortcut via Graph](https://techcommunity.microsoft.com/discussions/sharepointdev/onedrive-shortcut-via-graph/3251924)
- App-Ordner: [Graph: App folder in OneDrive and SharePoint](https://learn.microsoft.com/en-us/graph/onedrive-sharepoint-appfolder) · [Q&A: AppFolder für Business](https://learn.microsoft.com/en-us/answers/questions/1418013/when-is-api-permission-(delegated)-files-readwrite)
- Anheften: [Pinnable task pane](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/pinnable-taskpane)
- Grenzen: [Limits for Outlook add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/limits-for-activation-and-javascript-api-for-outlook-add-ins) · [SharePoint limits](https://learn.microsoft.com/en-us/office365/servicedescriptions/sharepoint-online-service-description/sharepoint-online-limits) · [Upload small files](https://learn.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0) · [createUploadSession](https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession?view=graph-rest-1.0)
- .eml im neuen Outlook: [Microsoft Support](https://support.microsoft.com/en-us/outlook/mail/open-eml-msg-and-oft-files-in-new-outlook-and-outlook-on-the-web)
- Zeitplan klassisches Outlook: [PCWorld](https://www.pcworld.com/article/3082363/microsoft-extends-support-for-the-classic-outlook-app-again.html) · [Tech Community](https://techcommunity.microsoft.com/discussions/microsoft-365/outlook-classic-support-until-at-least-2029/4081174)
- Hosting/App-Katalog: [Publish to SharePoint app catalog](https://learn.microsoft.com/en-us/office/dev/add-ins/publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog)
- WebView2/ONNX: [ONNX Runtime Web](https://onnxruntime.ai/docs/tutorials/web/) · [Issue #1913 WebView2-Speicher](https://github.com/OfficeDev/office-js/issues/1913)

Hinweis: `learn.microsoft.com` war für direkte Seitenabrufe gesperrt; die Aussagen stützen sich auf
Suchergebnis-Auszüge und die Spiegel der Microsoft-Doku auf GitHub. Vor der Umsetzung stichprobenartig im
Original gegenprüfen.

