# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

MailDrop is a **Visual Studio Tools for Office (VSTO) Outlook Add-in** written in Visual Basic .NET that files emails and attachments into project folder structures. It uses a BERT model (via ONNX) to suggest mail drop configuration.

- **Language**: Visual Basic .NET, .NET Framework 4.7.2, VSTO 4.0
- **Host**: Microsoft Outlook
- **Build tool**: MSBuild / Visual Studio 2022

## Features

- Outlook ribbon button MailDrop opens a right-side task pane.
- The add-in reacts to selection changes and enables editing only when exactly one mail is selected.
- Session data is prepared from the selected mail metadata (Betreff, Absender, AbsenderDomain, Empfaenger, Datum).
- ProjektPfad offers recent project folders for the current user (from SQLite history) plus anderes... via folder picker.
- ProjektstrukturPfad is selected from a TreeView that is built dynamically from the selected ProjektPfad.
- Titel and Absender (kurz) are editable text fields used by placeholder resolution.
- Ablageordner and msg Dateiname use a Schema/Aufgeloest/Feld workflow:
	- On focus: show schema text for editing.
	- On focus loss: resolve placeholders and show the resolved result.
- Input validation checks required values, invalid path/file characters, and path length limits.
- Optional attachment filing saves all attachments; long attachment names can be adjusted via rename dialog.
- On OK, the add-in creates the target folder, saves the mail as .msg, optionally saves attachments, and stores the session in SQLite.
- A SuggestionEngine with ONNX embeddings is initialized for project path suggestion.
- During session preparation, the engine precomputes feature distance lists between current mail/session and historical records, then scores a suggested ProjektPfad.
- Suggestion-relevant features are split into two groups: fixed features (Betreff, Datum, AbsenderDomain, Absender, AusfueBenutzer) are computed once per mail selection via dedicated per-feature subs called from PrepareSession(); mutable features (Titel, AblageordnerAufgeloest, ProjektPfad, ProjektstrukturPfad) are recomputed via dedicated per-feature subs triggered from the corresponding Session property setters and are zero-initialized at engine construction.
- Distance list allocation and initial computation are owned by CalculateInitialFeatureDistances(session) on SuggestionEngine, called once from PrepareSession() immediately after New(). New() handles only engine infrastructure (historical records, EmbeddingService, historical embeddings).

## Current workspace layout:

MailDrop/
|- app.config
|- CLAUDE.md
|- MailDrop.sln
|- MailDrop.vbproj
|- model.onnx
|- packages.config
|- packages.config.old.20251107221305
|- packages.config.old.20251111163110
|- ThisAddIn.Designer.vb
|- ThisAddIn.Designer.xml
|- ThisAddIn.vb
|- vocab.txt
|- Core/
|  |- InputChecker.vb
|  |- Session.vb
|  |- SuggestionEngine.vb
|- Helpers/
|  |- DatabaseUtils.vb
|  |- DirectoryTreeHelper.vb
|  |- MailUtils.vb
|  |- PathShortenerConverter.vb
|- Models/
|  |- model.onnx
|  |- vocab.txt
|- My Project/
|  |- AssemblyInfo.vb
|  |- Resources.Designer.vb
|  |- Resources.resx
|  |- Settings.Designer.vb
|  |- Settings.settings
|- Services/
|  |- EmbeddingService.vb
|- TestDirectory/
|  |- P-23002/
|  |- P-23003_Kita/
|  |- P-23004_Modehaus/
|- UI/
	|- AttachmentRenameDialog.xaml
	|- AttachmentRenameDialog.xaml.vb
	|- InfoPopup.xaml
	|- InfoPopup.xaml.vb
	|- MailDropRibbon.vb
	|- MailDropRibbon.xml
	|- MailDropWpfHostControl.vb
	|- MailDropWpfTaskPane.xaml
	|- MailDropWpfTaskPane.xaml.vb

## Build and run (local)

- Open MailDrop.sln in Visual Studio 2022.
- Ensure NuGet packages are restored (packages.config based project).
- Build Debug|AnyCPU for normal local development.
- Start debugging from Visual Studio; Outlook is the host process for the VSTO add-in.
- Requirement: Outlook desktop with VSTO runtime available.

## Runtime data and paths

- Session history database: %APPDATA%/MailDrop/sessions.db
- Database manager startup: ThisAddIn_Startup creates SessionDatabaseManager.
- ONNX model files are loaded from output folder path Models/model.onnx and Models/vocab.txt.
- The root-level model.onnx and vocab.txt are source artifacts; runtime inference uses the files under Models/.

## Architecture and control flow

- Entry point: ThisAddIn_Startup in ThisAddIn.vb.
- Ribbon action: MailDropRibbon -> Globals.ThisAddIn.MailAblegen_Click.
- Task pane creation: MailDropWpfHostControl hosts MailDropWpfTaskPane.
- Selection updates: Explorer.SelectionChange -> MailSelected().
- Session preparation: Session.PrepareSession() -> Reset, mail metadata read, recent project paths load, SuggestionEngine init, feature distance precompute, suggested ProjektPfad apply (if valid).
- Save flow on OK: Session.ProcessSession() -> InputChecker -> create folder -> save .msg -> optional attachments -> persist SessionRecord -> reset session.

## Key files for first orientation

- ThisAddIn.vb: add-in lifecycle, Outlook selection event wiring, task pane lifecycle.
- UI/MailDropRibbon.vb and UI/MailDropRibbon.xml: ribbon integration and button callback.
- UI/MailDropWpfTaskPane.xaml and UI/MailDropWpfTaskPane.xaml.vb: primary UI and event handling.
- Core/Session.vb: central state object and business flow.
- Core/InputChecker.vb: input validation and attachment rename dialog path handling.
- Helpers/MailUtils.vb: Outlook MailItem metadata read/save operations.
- Helpers/DatabaseUtils.vb: SQLite persistence (SessionRecord) and retrieval of recent project paths.
- Services/EmbeddingService.vb and Core/SuggestionEngine.vb: ONNX embedding and suggestion logic.

## Placeholder reference

Supported placeholders for Ablageordner and msg Dateiname:

- [Titel]
- [Absender]
- [Absender-Domain]
- [Empfaenger]
- [Empfaenger (kurz)]
- [Betreff]
- [Datum]
- [Datum (formatiert)]
- [Absender (kurz)]

Note:
- [Empfaenger (kurz)] currently resolves to the same value as Empfaenger.

## Validation and constraints

- ProjektPfad must exist.
- ProjektstrukturPfad must be selected and must exist under ProjektPfad.
- Invalid file/path characters are rejected.
- Path length limit checks currently use 255 characters.
- Attachment save can prompt for rename when resulting path is too long.

## Current caveats and implementation notes

- SuggestionEngine uses weighted scoring across Betreff (semantic cosine similarity), Datum (normalized date similarity), AbsenderDomain (categorical match), Absender (categorical match), AusfueBenutzer (categorical match), Titel (text similarity), Ablageordner (text similarity), ProjektPfad (categorical match), and ProjektstrukturPfad (categorical match).
- Feature weights are defined per feature directly in SuggestProjektPfad in Core/SuggestionEngine.vb and are applied independently during score aggregation.
- Distance lists are initialized with zeros at construction (length = historical record count). Fixed-feature subs overwrite them in PrepareSession(); mutable-feature subs overwrite them incrementally as the user edits fields.
- Historical Betreff embeddings are created on demand in memory if missing in persisted records.
- Current feature weighting is heuristic constants in Core/SuggestionEngine.vb and may need tuning with real usage data.
- In Core/InputChecker.vb, the Path.Combine call for ablageOrdnerPfad combines projektPfad and an already-combined projektstrukturPfad, which can duplicate path segments.
- Text encoding/comments show mixed umlaut encoding artifacts in several files; prefer preserving existing file encoding unless intentionally normalizing.

## Conventions for Claude Code contributions

- Keep domain terms and UI labels in German to match existing code and user-facing text.
- Favor minimal, targeted edits; do not refactor broad areas unless needed for the requested change.
- Preserve the Session-based flow (PrepareSession/ProcessSession) and avoid introducing parallel state objects.
- For UI behavior changes, keep data-binding centered on Session properties and existing control names.
- For persistence changes, keep SessionRecord schema and SQLite migration impact explicit.
- When touching SuggestionEngine, keep behavior behind existing calls to avoid breaking current filing flow.
- Whenever code is changed, also review and update this CLAUDE.md so implementation notes, behavior descriptions, and caveats stay in sync.

## Quick verification checklist after changes

- Build succeeds in Visual Studio.
- Ribbon button opens the task pane.
- Exactly one selected mail enables editing; other selections disable editing.
- Selecting ProjektPfad refreshes Projektstruktur TreeView.
- Placeholder fields resolve correctly on focus loss.
- OK saves .msg to expected folder.
- Optional attachment save works and handles long names.
- Session row is written to SQLite database.



