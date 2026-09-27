# MailDrop Web – technischer Prototyp

Dieser Ordner enthält den **Prototyp** des MailDrop-Web-Add-ins (Outlook-Add-in auf Basis von
Office.js). Er legt noch keine echten Ablagen an, sondern prüft die offenen Punkte aus
[`docs/Web-Add-in-Umbau.md`](../docs/Web-Add-in-Umbau.md), Abschnitt 10, in eurer echten
Microsoft-365-Umgebung:

| Nr. | Prüfung | Klärt |
|---|---|---|
| 1 | Umgebung und Requirement Sets | Outlook-Variante, Mailbox 1.8/1.14, NAA verfügbar? |
| 2 | Anmeldung (NAA) | Funktioniert die Anmeldung über das Outlook-Konto? |
| 3 | Aktuelle Mail | Metadaten lesbar, freigegebenes Postfach erkannt? |
| 4 | .eml-Export | `getAsFileAsync` auf euren Outlook-Versionen, Graph-Rückfallebene, S/MIME/Vertraulichkeit |
| 5 | Anhänge | Abruf aller Anhänge nacheinander |
| 6 | Projektbibliotheken finden | Werden OneDrive-Verknüpfungen über Graph gelistet? Rückfallebene gefolgte Sites |
| 7 | Speicherort Verlauf | App-Ordner oder Rückfallebene `Apps/MailDrop` |
| 8 | Test-Upload (nur manuell) | .eml in eine gewählte Projektbibliothek hochladen |
| 9 | Vorschlagsmodell | Laden/Rechnen des Modells im Aufgabenbereich, Ladezeit |

Prüfung 8 **schreibt** eine Testdatei in den Ordner `MailDrop-Prototyp-Test` der gewählten
Bibliothek und läuft deshalb nur per Klick, nie bei „Alle ausführen“. Prüfung 7 schreibt eine
kleine Testdatei in den App-Ordner bzw. `Apps/MailDrop` des eigenen OneDrive.

## Voraussetzungen

1. **App-Registrierung durch die IT** (siehe `docs/Web-Add-in-Umbau.md`, Abschnitt 7):
   - Plattform **Single-Page-Anwendung** mit den Umleitungs-URIs
     `brk-multihub://localhost:3000` und `https://localhost:3000/taskpane.html`
   - Delegierte Berechtigungen `User.Read`, `Files.ReadWrite.All`, `Files.ReadWrite.AppFolder`,
     `Mail.Read`, `Mail.Read.Shared`, `Sites.Read.All` (letztere nur für die Rückfallebene
     „gefolgte Sites“ in Prüfung 6), jeweils mit Administrator-Zustimmung
2. **Node.js 20 oder neuer** auf dem Test-PC
3. Hochladen eigener Add-ins in Outlook darf nicht per Richtlinie gesperrt sein – sonst muss die
   IT das Manifest über das Admin-Center an dich verteilen

## Starten

```powershell
cd web
npm install
copy .env.example .env.local        # Client- und Tenant-ID aus der App-Registrierung eintragen
npx office-addin-dev-certs install   # einmalig: Entwicklerzertifikat für https://localhost vertrauen
npm run dev                          # startet https://localhost:3000
```

Dann das Add-in in Outlook hinzufügen:

- **Neues Outlook / Outlook im Browser:** eine Mail öffnen → „Apps“ bzw. „…“ → „Add-Ins
  abrufen“ → „Meine Add-Ins“ → „Benutzerdefiniertes Add-In hinzufügen“ → „Aus Datei“ →
  `web/manifest.xml`
- **Klassisches Outlook:** Datei → „Add-Ins verwalten“ öffnet dieselbe Seite im Browser; dort wie
  oben

Danach eine Mail **mit Anhang** auswählen, im Menüband „Mail ablegen (Prototyp)“ klicken,
Aufgabenbereich anheften (Stecknadel), **„Alle ausführen“**, anschließend Prüfung 8 mit einer
Bibliothek aus der Liste, dann **„Bericht kopieren“** und den Bericht zurückmelden.

Aussagekräftig wird es, wenn die Prüfungen in **jeder** eingesetzten Variante laufen: klassisches
Outlook, neues Outlook, Browser – und einmal mit einer Mail aus einem **freigegebenen Postfach**
sowie mit einer **verschlüsselten/als vertraulich gekennzeichneten** Mail, falls es die gibt.

## Befehle

| Befehl | Zweck |
|---|---|
| `npm run dev` | Entwicklungsserver auf https://localhost:3000 (kopiert vorher das Modell aus `Models/`) |
| `npm run build` | Typprüfung + Produktions-Build nach `web/dist` |
| `npm test` | Tests (Tokenizer, Modell mit ONNX Runtime Web in Node) |
| `npx office-addin-manifest validate manifest.xml` | Manifest bei Microsoft validieren (braucht Internet) |

## Aufbau

- `manifest.xml` – Add-in-Beschreibung (Button im Lesebereich und geöffneten Mail-Fenster, anheftbar, freigegebene Postfächer)
- `src/taskpane.html`, `src/taskpane.ts` – Prüfoberfläche
- `src/checks.ts` – die neun Prüfungen
- `src/auth.ts` – Anmeldung (NAA, Rückfallebene Popup)
- `src/graph.ts`, `src/mail.ts` – Graph- und Office.js-Zugriffe
- `src/embedding/` – Tokenizer (Unigram für das deutschfähige Modell, WordPiece als Rückfall für das alte englische) und Modell; Tests lesen mit `MAILDROP_MODELS_DIR` auch einen anderen Modellordner
- `scripts/copy-models.mjs` – kopiert `Models/model.onnx` und `vocab.txt` nach `public/models` (nicht eingecheckt)
