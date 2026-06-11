# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

MailDrop is a **Visual Studio Tools for Office (VSTO) Outlook Add-in** written in Visual Basic .NET that files emails and attachments into project folder structures. It uses a BERT model (via ONNX) to suggest the best matching project folder based on the email subject.

- **Language**: Visual Basic .NET, .NET Framework 4.7.2, VSTO 4.0
- **Host**: Microsoft Outlook
- **Build tool**: MSBuild / Visual Studio 2022

## Architecture

### End-to-end flow

`Explorer_SelectionChange` → `MailSelected()` — silent no-op until the ribbon button is clicked for the first time (`GetWpfTaskPane()` returns `Nothing` while the task pane hasn't been created yet).

Once the pane exists: `PrepareSession()` reads mail metadata, loads recent project directories, and runs `SuggestionEngine` to pre-select a project. The user picks a project and sub-folder, edits the `Ablageordner`/`MsgDateiname` fields, and clicks OK. `ProcessSession()` then validates paths (`InputChecker.CheckInput()`), creates the folder, saves the `.msg`, optionally saves attachments, persists a `SessionRecord` to SQLite, and resets.

### Session — the central state object (`Core/Session.vb`)

Implements `INotifyPropertyChanged`, set as `DataContext` of the task pane — all WPF bindings read/write it directly.

**Schema/Aufgeloest/Feld triad** — `Ablageordner` and `MsgDateiname` each have three related properties: `*Schema` (raw template, e.g. `[Datum] [Titel]`), `*Aufgeloest` (read-only resolved value), and `*Feld` (what the TextBox shows). GotFocus swaps `Feld ← Schema`; LostFocus writes `Schema ← Feld`, recomputes `Aufgeloest`, then swaps `Feld ← Aufgeloest`. Orchestrated by `BeginAblageordnerEdit()` / `EndAblageordnerEdit()`.

**`ToSessionRecord()`** copies properties from `Session` to `SessionRecord` by name-matching via reflection — both types must stay in sync.

**Path construction** (`InputChecker.CheckInput()`):
```
projektstrukturPfad (absolute) = Path.Combine(ProjektPfad, Session.ProjektstrukturPfad)
ablageOrdnerPfad               = Path.Combine(projektstrukturPfad, AblageordnerAufgeloest)
msgZielPfad                    = Path.Combine(ablageOrdnerPfad, MsgDateinameAufgeloest)
```
`Session.ProjektstrukturPfad` holds a **relative** path (relative to `ProjektPfad`); it is set from `DirectoryNode.RelativePath` when the user clicks a tree node.

### Placeholder system (`Core/Session.vb` — `ReplacePlaceholders`)

Templates in `AblageordnerSchema` and `MsgDateinameSchema` may contain:

| Placeholder | Source property |
|---|---|
| `[Titel]` | `Titel` (user-editable) |
| `[Betreff]` | `Betreff` (email subject) |
| `[Absender]` | `Absender` (full sender name) |
| `[Absender (kurz)]` | `AbsenderKurz` (user-editable in UI) |
| `[Absender-Domain]` | `AbsenderDomain` (derived from SMTP address) |
| `[Empfänger]` / `[Empfänger (kurz)]` | `Empfaenger` |
| `[Datum]` | `Datum` formatted as `yyyy-MM-dd` |
| `[Datum (formatiert)]` | `DatumFormatiert` (formatted as `yyyyMMdd`) |

Note: `AbsenderKurz` is not auto-filled by `ReadMailMeta` — the user types it directly. `AbsenderDomain` is parsed from the SMTP address only when `SenderEmailType = "SMTP"` and the address contains `@`.

### SuggestionEngine (`Core/SuggestionEngine.vb`)

Scores every historical `SessionRecord` and returns the `ProjektPfad` of the best match. Current weight vector:

```
0.4 × SemanticFeature   (cosine similarity of BERT-embedded subjects)
0.2 × NumericalFeature  (date distance — stub, returns 0)
0.1 × CategorialFeature (sender domain — stub, returns 0)
0.1 × CategorialFeature (sender — stub, returns 0)
0.2 × CategorialFeature (user — stub, returns 0)
```

The comment `'HIER GEHTS WEITER` at line 130 marks where the unimplemented features should be filled in next.

`EmbedBetreff(session)` is called during `PrepareSession()` to store the current email's embedding on the `Session` object for later persistence. `SuggestProjektPfad(session)` recomputes embedding and distances immediately from `EnginesHistoricalSessionRecords`.

### EmbeddingService (`Services/EmbeddingService.vb`)

Loads `model.onnx` (384-dimensional BERT) and `vocab.txt` at construction. `GenerateEmbedding(text)`:
1. Tokenizes with `BertTokenizer`
2. Builds `input_ids`, `attention_mask`, `token_type_ids` tensors (shape `[1, seqLen]`)
3. Runs ONNX inference
4. Mean-pools the token dimension over the `[1, seqLen, 384]` output
5. Applies L2 normalization

### Database (`Helpers/DatabaseUtils.vb`)

`SessionDatabaseManager` manages `%APPDATA%\MailDrop\sessions.db`. The `dbPath` field is assigned both as a field initializer and again in the constructor — redundant but harmless. Key points:

- `SaveSessionRecord()` builds its `INSERT` dynamically via reflection over `SessionRecord` properties, skipping `ID` (AUTOINCREMENT). The `BetreffEmbedded` `Single()` array is serialized with `DatabaseUtils.FloatsToBytes()` (`Buffer.BlockCopy`).
- `GetAllSessionRecords()` deserializes `BetreffEmbedded` back with `DatabaseUtils.BytesToFloats()`.
- `GetLastProjektVerzeichnisseForUser()` queries `SELECT DISTINCT ProjektPfad ... ORDER BY AusfueDatum DESC LIMIT 10`, then the VB loop exits early once 4 results are collected. Caution: ordering by `AusfueDatum` while selecting only `ProjektPfad` (DISTINCT) means SQLite picks an arbitrary row per distinct path before ordering, so the "most recent" ordering is not strictly guaranteed.
- `SessionRecord` is the plain DTO. `EncodedSessionRecord` and the commented-out `EncodedSessions` table are an abandoned earlier ML approach — do not resurrect.

### UI (`UI/`)

**WinForms/WPF bridge**: `MailDropWpfHostControl` (WinForms `UserControl`) wraps an `ElementHost` whose `Child` is `MailDropWpfTaskPane` (WPF `UserControl`). This double wrapping is required because VSTO custom task panes only accept WinForms controls.

**`MailDropWpfTaskPane.xaml`** bindings summary:
- `ListBox1` ← `ProjektVerzeichnisse`; items display via `PathShortenerConverter` (`C:\Drive\...\Folder`) with full path as tooltip
- `TreeView1` ← `TreeViewData`; `HierarchicalDataTemplate` over `DirectoryNode.Children`; first two levels auto-expand (`IsExpanded` bound via `DataTrigger`)
- `TextBoxTitel` ↔ `Titel`
- `TextBoxAbsenderKurz` ↔ `AbsenderKurz`
- `TextBoxAblageordner` ↔ `AblageordnerFeld` (with GotFocus/LostFocus schema-swap)
- `TextBoxMsgDateiname` ↔ `MsgDateinameFeld` (with GotFocus/LostFocus schema-swap)
- `CheckBoxAnhaenge` ↔ `AnhaengeAblegen`

**Ribbon**: `MailDropRibbon.vb` implements `IRibbonExtensibility`. The XML (`MailDropRibbon.xml`) is embedded as a manifest resource (`MailDrop.MailDropRibbon.xml`). The button appears in `TabMail` (the Outlook "Start" tab). Clicking delegates to `Globals.ThisAddIn.MailAblegen_Click()`, which creates the task pane on first use (docked right, width 1000 px) and then calls `MailSelected()`. The task pane is a **singleton** — once created it is only ever shown or hidden via `taskPane.Visible`; it is never recreated. `HideTaskPane()` (called by the Cancel button) sets `taskPane.Visible = False` without destroying it.

**`GetWpfTaskPane()` in `ThisAddIn.vb`** navigates: `taskPane.Control` (WinForms `UserControl`) → `.Controls(0)` (`ElementHost`) → `.Child` (`MailDropWpfTaskPane`).

**Shared properties on `ThisAddIn`**: `DbDirectory` (`%APPDATA%\MailDrop`) is used by `SessionDatabaseManager` to locate the database. `DbPath` (the full `.db` path) is also declared as a shared property but is not read externally — `SessionDatabaseManager` recomputes the path itself. `CurrentDatabaseManager` is the global singleton database manager, accessible anywhere via `ThisAddIn.CurrentDatabaseManager`.

### InputChecker (`Core/InputChecker.vb`)

A `Module` (not a class). `CheckInput(session)` validates in order: ProjektPfad exists → ProjektstrukturPfad non-empty and exists → AblageordnerAufgeloest valid chars and ≤255 chars → MsgDateinameAufgeloest valid chars and ≤255 chars → each attachment path valid (triggers `AttachmentRenameDialog` when an attachment path exceeds 255 chars, and skips the attachment if the user cancels). Returns a `CheckedInputResult` with all resolved paths or a non-empty `ErrorMessage`.

### DirectoryTreeHelper (`Helpers/DirectoryTreeHelper.vb`)

`BuildDirectoryTree(projektPfad)` returns an `ObservableCollection(Of DirectoryNode)` containing the direct children of `projektPfad`, each recursively populated. `DirectoryNode.RelativePath` is the path **relative to `projektPfad`** and is what gets assigned to `Session.ProjektstrukturPfad` on tree selection. Nodes at depth ≤ 2 have `IsExpanded = True`.

## German naming conventions

All domain identifiers, comments, and UI labels are in German.

| German | Meaning |
|---|---|
| Betreff | Email subject |
| Absender | Sender |
| Empfaenger | Recipient |
| Ablageordner | Filing/storage folder |
| ProjektPfad | Project root path (absolute) |
| ProjektstrukturPfad | Sub-folder path within project (relative) |
| Anhang / Anhaenge | Attachment(s) |
| Ausfue* | Execution context (AusfueDatum = when filed, AusfueBenutzer = who filed) |
| Schema | Template string with placeholders |
| Aufgeloest | Resolved string (placeholders substituted) |
| Feld | Current value shown in the UI text box |
