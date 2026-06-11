# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

MailDrop is a **Visual Studio Tools for Office (VSTO) Outlook Add-in** written in Visual Basic .NET that intelligently files emails and attachments into project folder structures. It uses BERT embeddings (via ONNX) to suggest the best project folder based on the email subject.

- **Language**: Visual Basic .NET
- **Framework**: .NET Framework 4.7.2
- **Host application**: Microsoft Outlook (VSTO 4.0)
- **Build tool**: MSBuild / Visual Studio 2022

## Build

Build from the solution root using MSBuild:

```
msbuild MailDrop.sln /p:Configuration=Debug
msbuild MailDrop.sln /p:Configuration=Release
```

Restore NuGet packages before the first build (Visual Studio does this automatically; on the command line use `nuget restore MailDrop.sln`).

Output lands in `bin\Debug\` or `bin\Release\`. The ONNX model and vocabulary file (`Models\model.onnx`, `Models\vocab.txt`) are copied to `bin\<Config>\Models\` by the project at build time — `EmbeddingService` resolves their path relative to the executing assembly.

## Testing

There is no automated test suite. Manual integration testing is done by running the add-in inside Outlook (F5 in Visual Studio starts Outlook with the add-in loaded). The `TestDirectory/` folder contains a realistic sample project tree for testing folder navigation and suggestion features.

## Architecture

### Data flow

1. User selects an email in Outlook → `Explorer_SelectionChange` fires in `ThisAddIn.vb`
2. `MailDropWpfTaskPane.SingleMailSelected()` checks the selection
3. `Session.PrepareSession()` calls `MailUtils.ReadMailMeta()` to populate mail properties, then `SuggestionEngine.SuggestProjektPfad()` to propose a project folder
4. User adjusts the folder/filename in the task pane and clicks OK
5. `Session.ProcessSession()` validates paths via `InputChecker`, saves the `.msg` file and attachments, then persists a `SessionRecord` (including the BERT embedding of the subject) to SQLite

### Key classes

**`Core/Session.vb`** — Central state object and effective ViewModel. Implements `INotifyPropertyChanged` so all WPF controls bind to it directly. Responsible for:
- Holding all mail metadata and user-editable fields
- Resolving placeholder templates (see *Placeholders* below)
- Orchestrating `PrepareSession()` and `ProcessSession()`

**`Core/SuggestionEngine.vb`** — Scores each historical `SessionRecord` against the current email using weighted features and returns the best-matching project path. Currently only the `SemanticFeature` (subject cosine similarity) is implemented; `NumericalFeature` (date) and `CategorialFeature` (sender, domain, user) have TODO stubs returning 0.

**`Services/EmbeddingService.vb`** — Loads `model.onnx` (a 384-dimensional BERT model) and `vocab.txt`, tokenizes text, runs ONNX inference, applies mean-pooling across tokens, and returns an L2-normalized `Single()` vector.

**`Helpers/DatabaseUtils.vb`** — `SessionDatabaseManager` manages the SQLite database at `%APPDATA%\MailDrop\sessions.db`. The `BetreffEmbedded` column stores the embedding vector as a BLOB; `DatabaseUtils.FloatArrayToBytes` / `BytesToFloatArray` handle serialization. `GetLastProjektVerzeichnisseForUser()` returns the 3 most-recently-used project paths for the list box.

**`UI/MailDropWpfTaskPane.xaml(.vb)`** — WPF `UserControl` docked in Outlook's right task pane. `Me.DataContext = Session` wires up all bindings. The WinForms/WPF interop bridge is `MailDropWpfHostControl.vb` (a `UserControl` containing an `ElementHost`).

**`ThisAddIn.vb`** — Add-in entry point. Exposes `Shared Property CurrentDatabaseManager` so any module can reach the database without passing references around.

### Placeholder system

Path and filename templates use bracketed tokens resolved in `Session.ReplacePlaceholders()`:

| Placeholder | Source |
|---|---|
| `[Titel]` | User-editable title field |
| `[Betreff]` | Raw email subject |
| `[Absender]` | Full sender address |
| `[Absender (kurz)]` | Short form of sender |
| `[Absender-Domain]` | Domain part of sender |
| `[Empfänger]` | Recipient |
| `[Datum]` | Date as `yyyy-MM-dd` |
| `[Datum (formatiert)]` | Human-formatted date |

The convention is: `*Schema` = raw template string, `*Aufgeloest` = resolved read-only result, `*Feld` = the value currently displayed/edited in the UI text box.

### German naming

All domain identifiers, comments, and UI labels are in German. Key terms:

| German | English |
|---|---|
| Betreff | Email subject |
| Absender | Sender |
| Empfaenger | Recipient |
| Ablageordner | Storage/filing folder |
| ProjektPfad | Project root path |
| ProjektstrukturPfad | Sub-folder path within the project (relative) |
| Anhang/Anhaenge | Attachment(s) |
| Ausfuehrung/Ausfue* | Execution (e.g. AusfueDatum = execution date) |

## Active TODOs

- `SuggestionEngine.vb`: `NumericalFeature.CalculateDistances()` and `CategorialFeature.CaclulateDistances()` are stubs; the comment `'HIER GEHTS WEITER` marks where date/sender/domain/user scoring should be implemented next.
- `Session.vb`: The `PredictionEngine.TrainDecisionTreeModels()` call is commented out — an earlier ML approach that was replaced by the ONNX embedding approach.
- `DatabaseUtils.vb`: The `EncodedSessions` table and related encoding methods are commented out (abandoned alternative ML approach).
