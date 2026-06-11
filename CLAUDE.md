# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

MailDrop is a **Visual Studio Tools for Office (VSTO) Outlook Add-in** written in Visual Basic .NET that files emails and attachments into project folder structures. It uses a BERT model (via ONNX) to suggest the best matching project folder based on the email subject.

- **Language**: Visual Basic .NET, .NET Framework 4.7.2, VSTO 4.0
- **Host**: Microsoft Outlook
- **Build tool**: MSBuild / Visual Studio 2022

## Build

```
msbuild MailDrop.sln /p:Configuration=Debug
msbuild MailDrop.sln /p:Configuration=Release
```

Restore NuGet packages before the first build (`nuget restore MailDrop.sln`; Visual Studio does this automatically). Output lands in `bin\Debug\` or `bin\Release\`. The ONNX model files (`Models\model.onnx`, `Models\vocab.txt`) are copied into `bin\<Config>\Models\` at build time — `EmbeddingService` resolves them relative to the executing assembly.

## Testing

There is no automated test suite. Manual integration testing is done by pressing F5 in Visual Studio, which launches Outlook with the add-in loaded. The `TestDirectory/` folder contains a sample project tree for exercising folder navigation and suggestion features.

## Architecture

### End-to-end flow

1. User selects an email → `Explorer_SelectionChange` fires in `ThisAddIn.vb` and calls `MailSelected()`. If the task pane has never been opened, `GetWpfTaskPane()` returns `Nothing` and `MailSelected()` is a silent no-op — the add-in is lazy until the ribbon button is first clicked.
2. `MailDropWpfTaskPane.SingleMailSelected()` checks that exactly one `MailItem` is selected
3. `Session.PrepareSession()` runs: resets state, calls `MailUtils.ReadMailMeta()` to populate mail fields, calls `GetProjektVerzeichnisse()` to load the list-box, creates a `SuggestionEngine` and calls `SuggestProjektPfad()` to set a default project folder
4. User selects a project from the list-box (or picks "anderes..." for a folder browser), picks a sub-folder in the tree-view, edits the `Ablageordner` and `MsgDateiname` fields, and clicks OK
5. `Session.ProcessSession()`: validates with `InputChecker.CheckInput()`, creates the target folder, saves the `.msg` file via `MailUtils.SaveSelectedMailAsMsg()`, optionally saves attachments via `MailUtils.SaveMailAttachments()`, persists a `SessionRecord` to SQLite, then resets

### Session — the central state object (`Core/Session.vb`)

`Session` is both the data model and the effective ViewModel. It implements `INotifyPropertyChanged` and is set as `DataContext` of the task pane, so all WPF bindings read and write directly to it.

**The Schema/Aufgeloest/Feld triad** — for `Ablageordner` and `MsgDateiname` there are three related properties each:

- `*Schema` — the raw template string, e.g. `[Datum] [Titel]`
- `*Aufgeloest` — read-only; auto-recomputed by `ReplacePlaceholders()` whenever schema or any mail field changes
- `*Feld` — the string currently shown in the TextBox

On **GotFocus** the TextBox switches to `*Schema` (so the user edits the template). On **LostFocus** it writes back to `*Schema`, recomputes `*Aufgeloest`, and shows the resolved value again. This swap is orchestrated by `BeginAblageordnerEdit()` / `EndAblageordnerEdit()` and the equivalents for `MsgDateiname`.

**`ToSessionRecord()`** uses reflection to copy properties from `Session` to `SessionRecord` by matching property names — both types must stay in sync.

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
