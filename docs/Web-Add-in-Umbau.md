# MailDrop als Web-Add-in – Umbauskizze

Stand: 2026-09-27 · Status: **Konzeptskizze, nichts davon ist umgesetzt**

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
| **SQLite-Datei im App-Ordner des Nutzer-OneDrive** | ja | nein | **Empfohlen.** Jede App bekommt dort einen eigenen, privaten Ordner (Berechtigung `Files.ReadWrite.AppFolder`), MailDrop sieht nur diesen. Die Datenbank bleibt SQLite (im Browser z. B. über `sql.js`/`wa-sqlite`), das Schema kann weitgehend übernommen werden. Verlauf ist automatisch auf allen Geräten da – der heutige manuelle Export („Vorschlagsdaten exportieren“) entfällt. |
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

**Entscheidung:** Als Projekte werden die Dokumentbibliotheken angeboten, die der Nutzer per
**„Verknüpfung zu ‚Meine Dateien‘ hinzufügen“** in sein OneDrive eingebunden hat. Das Add-in liest diese
Verknüpfungen über Graph (Einträge im Stammordner des Nutzer-OneDrive, die auf ein anderes Laufwerk verweisen)
und bietet sie als ProjektPfad-Liste an; ganz oben die zuletzt genutzten (aus dem Verlauf).

- Kein Hub, kein Namensschema und keine Suche nötig – jeder sieht genau „seine“ Projekte.
- Neues Projekt verfügbar machen = in SharePoint einmal „Verknüpfung hinzufügen“ klicken.
- **Wichtig:** Nur Verknüpfungen sind über Graph sichtbar. Eine Bibliothek, die lediglich per „Synchronisieren“
  lokal eingebunden ist, kennt Graph nicht – Nutzer müssen also die Verknüpfung verwenden (die zusätzlich lokal
  synchronisiert sein darf).
- Gespeichert werden Laufwerk-ID und Element-ID der Bibliothek + Anzeigename statt eines Pfads.
- Rückfallebene: Site-Adresse einfügen.

### E5 – Format der abgelegten Mail

Office.js und Graph liefern eine Mail nur als **.eml** (MIME), **nicht als .msg**. Die .eml wird **genauso
abgelegt** wie heute die .msg: gleicher Ablageordner, Dateiname aus dem Dateinamen-Schema, nur mit Endung
`.eml` (Feld „msg Dateiname“ wird zu „Dateiname“). Anhänge können wie heute zusätzlich einzeln abgelegt werden.

- **Inhalt:** Die .eml ist die vollständige Mail inklusive Kopfzeilen, Text und **eingebetteter Anhänge** –
  fachlich gleichwertig zur .msg.
- **Beschaffung ohne Zusatzberechtigung:** `item.getAsFileAsync()` (Office.js, Requirement Set Mailbox 1.14)
  liefert die .eml direkt aus dem geöffneten Element. Alternative über Graph (`/messages/{id}/$value`) bräuchte
  zusätzlich die Berechtigung `Mail.Read` – vermeiden.
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
  384-dimensionales Modell, Mean Pooling, L2-Normalisierung). Im Browser mit ONNX Runtime Web bzw. einer
  fertigen Bibliothek (z. B. Transformers.js), die genau diese Schritte bereits mitbringt – kaum eigener Code.
- **Abgleich:** Einmalig Referenz-Embeddings für eine Liste typischer Betreffzeilen erzeugen und prüfen, dass
  die Browser-Variante dieselben Werte liefert (u. a. gleiche Einstellung zur Kleinschreibung im Tokenizer).
- **Offener Punkt Laufzeit:** ~90 MB Modell im Aufgabenbereich des (neuen) Outlook – Ladezeit beim ersten Start
  und Speicherverbrauch im Pilot messen; bei Bedarf eine quantisierte Variante (~¼ der Größe) testen und die
  Vorschlagsqualität vergleichen.

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
2. **Technischer Durchstich (Prototyp):** Anmeldung, aktuelle Mail lesen, eine Projekt-Site finden, Ordner
   anzeigen, .eml + Anhänge hochladen. Klärt die riskanten Punkte früh (Berechtigungen, Anmeldung, Upload).
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
8. Ist der **OneDrive-App-Ordner** (`Files.ReadWrite.AppFolder`) für den Verlauf zulässig?
9. Gibt es **Richtlinien für bedingten Zugriff** (verwaltete Geräte, MFA, Standorte), die das Add-in betreffen?

**Azure und Verteilung**

10. Darf eine **Azure Static Web App** im Tenant betrieben werden? Welche Subscription/Ressourcengruppe, wer ist Besitzer?
11. Eigene Adresse (z. B. `maildrop.firma.de`) gewünscht?
12. Wer verteilt das Add-in im **Admin-Center (Integrierte Apps)**, an welche Gruppen?
13. Welche Outlook-Varianten sind im Einsatz (klassisch/neu/Web)? Im klassischen Outlook: sind **Outlook-Version** und **WebView2** aktuell genug für Web-Add-ins?
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
- [ ] Delegierte Graph-Berechtigungen eintragen: `User.Read`, `Files.ReadWrite.AppFolder`, `Files.ReadWrite.All`.
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
- **Erster Start** mit ~90 MB Modell – ggf. kleineres Modell nötig.
- **Vorschlagsqualität** nach Portierung – nur durch Abgleich mit der .NET-Version auf denselben Daten nachweisbar.
- **Graph-Drosselung** bei sehr großen Bibliotheken/Ordnerbäumen – durch ebenenweises Laden entschärft.
- **Office.js-Funktionsumfang** (Anheften, Schließen der Pane per Code, Mehrfachauswahl) hängt von den unterstützten Requirement Sets der eingesetzten Outlook-Versionen ab – vor Phase 2 prüfen.
