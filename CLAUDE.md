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
- ProjektstrukturPfad TreeView supports creating, deleting, and renaming folders via context menu actions (`Neuer Ordner`, `Loeschen`, `Umbenennen`), and also supports creating folders via a `Neuen Ordner` button; if a node is selected, operations target that node, otherwise creation defaults to ProjektPfad.
- Titel and Absender (kurz) are editable text fields used by placeholder resolution.
- Ablageordner and msg Dateiname use a Schema/Aufgeloest/Feld workflow:
	- On focus: show schema text for editing.
	- On focus loss: resolve placeholders and show the resolved result.
	- When suggestions set schema values during cascade, Feld values are updated immediately to the resolved result so users can see suggestions without focusing the field.
- Input validation checks required values, invalid path/file characters, and path length limits.
- Optional attachment filing saves all attachments; long attachment names can be adjusted via rename dialog.
- On OK, the add-in creates the target folder, saves the mail as .msg, optionally saves attachments, and stores the session in SQLite.
- A shared SuggestionEngine with ONNX embeddings is used for project path suggestion.
- The shared SuggestionEngine is lazy-loaded and preloaded in the background shortly after Outlook startup to reduce first task pane open latency.
- During session preparation, the engine precomputes feature distance lists between current mail/session and historical records, then scores a suggested ProjektPfad.
- Suggestion-relevant features are split into two groups: fixed features (Betreff, Datum, AbsenderDomain, Absender, AusfueBenutzer) are computed once per mail selection via dedicated per-feature subs called from PrepareSession(); mutable features (Titel, AblageordnerAufgeloest, ProjektPfad, ProjektstrukturPfad) are recomputed via dedicated per-feature subs triggered from the corresponding Session property setters and are zero-initialized at engine construction.
- Distance list allocation and initial computation are owned by CalculateInitialFeatureDistances(session) on SuggestionEngine, called once from PrepareSession() immediately after New(). New() handles only engine infrastructure (historical records); EmbeddingService is created lazily on first embedding generation.
- SuggestionEngine exposes SuggestProjektPfad, SuggestProjektstrukturPfad, SuggestTitel, and SuggestAblageordnerSchema, all backed by a shared FindBestRecord helper that scores all historical records and returns the best match for the requested field. SuggestAblageordnerSchema and SuggestMsgDateinameSchema fall back to suggesting AblageordnerAufgeloest / MsgDateinameAufgeloest (the stored resolved values as literals) when no schema templates have been stored yet, to avoid empty fields on first use.
- Internally, SuggestionEngine keeps the scoring loop inside FindBestRecordByField (with an optional non-empty-field filter switch) and avoids Func-signature overloading to prevent ambiguous lambda resolution in VB.
- For tracing/logging, SuggestionEngine passes explicit suggestion labels (e.g., ProjektPfad, MsgDateinameSchema) into the shared FindBestRecord helper instead of relying on lambda Method.Name.
- Session exposes SuggestXxx(value) methods that set the field value and then set the corresponding IsSuggestedXxx boolean flag to True. Cascade is driven by property setters: ProjektPfad → ProjektstrukturPfad → Titel → AbsenderKurz → AblageordnerSchema → MsgDateinameSchema → AnhaengeAblegen. Session also exposes IsSuggestedProjektPfad, IsSuggestedProjektstrukturPfad, IsSuggestedTitel, IsSuggestedAbsenderKurz, IsSuggestedAblageordnerSchema, IsSuggestedMsgDateinameSchema, IsSuggestedAnhaengeAblegen as full properties with PropertyChanged notification.
- MailDropWpfTaskPane.Session_PropertyChanged reacts to IsSuggestedXxx becoming True by showing the ✦ sparkle next to the field and (for ProjektstrukturPfad) dispatching a programmatic TreeView node selection via Dispatcher.BeginInvoke. The sparkle fades out when the user focuses/interacts with the field, which also sets the corresponding IsSuggestedXxx flag back to False (including AbsenderKurz, msg Dateiname and Anhänge ablegen).
- A private _applyingTreeViewSuggestion flag guards TreeView1_SelectedItemChanged against treating the engine's programmatic node selection as a user interaction (which would reset IsSuggestedProjektstrukturPfad and hide the sparkle immediately).
- The info popup (help button `i`) documents both placeholder usage and how users can transfer SuggestionRecords by copying `%APPDATA%/MailDrop/sessions.db` to another machine and placing it at the same path.

## Current workspace layout:

MailDrop/
|- app.config
|- CLAUDE.md
|- CONTRIBUTING.md
|- README.md
|- README.en.md
|- MailDrop.sln
|- MailDrop.vbproj
|- model.onnx
|- packages.config
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
- SQLite schema versioning uses `PRAGMA user_version`; `SessionDatabaseManager.CurrentSchemaVersion` defines the expected schema and `ApplyMigrations(...)` upgrades older databases.
- Database manager initialization is lazy: ThisAddIn.CurrentDatabaseManager creates SessionDatabaseManager on first access (not in ThisAddIn_Startup).
- ONNX model files are loaded from output folder path Models/model.onnx and Models/vocab.txt.
- The root-level model.onnx and vocab.txt are source artifacts; runtime inference uses the files under Models/.

## Branching model

- `productive`: stable branch intended for production-ready releases.
- `development`: integration branch for ongoing development changes.
- Feature branches (for example `feature/...`) should merge into `development`; release-ready states can then be promoted into `productive`.
- Repository governance details (review gates and merge flow) are documented in CONTRIBUTING.md.
- GitHub default branch should be `productive` (if still `master`, switch it in repository settings).

## Architecture and control flow

- Entry point: ThisAddIn_Startup in ThisAddIn.vb.
- Ribbon action: MailDropRibbon -> Globals.ThisAddIn.MailAblegen_Click.
- Task pane creation: MailDropWpfHostControl hosts MailDropWpfTaskPane.
- Task pane default width is set during first creation in ThisAddIn.MailAblegen_Click (currently 500 px) and docked right.
- Selection updates: Explorer.SelectionChange -> MailSelected().
- Session preparation: Session.PrepareSession() -> Reset, mail metadata read, recent project paths load, shared SuggestionEngine access, feature distance precompute, suggested ProjektPfad apply (if valid).
- Save flow on OK: Session.ProcessSession() -> InputChecker -> create folder -> save .msg -> optional attachments -> persist SessionRecord -> count check (every 50th record triggers RecalculateWeightsFromHistory in Task.Run on shared engine) -> reset session.
- SuggestionEngine pre-load: ThisAddIn_Startup fires PreloadSharedInstanceInBackground (SharedEngineLazy). Shared singleton loads historical records + computed weights from DB on first access.
- Weight recalculation: every 50th ProcessSession (recordCount Mod 50 = 0), Task.Run calls RecalculateWeightsFromHistory() on the shared engine. This reloads all records from DB, computes pairwise Pearson correlation per target field, persists results to ComputedWeights SQLite table, and replaces _computedWeights in memory. Non-milestone sessions skip this step.
- Feature weights are data-driven per target field when sufficient history exists. Weights are computed via pairwise Pearson correlation between feature similarity vectors and a binary label (same/different target value for each record pair), clipped to [0, ∞) and normalized to sum 1. Stored per target field in the ComputedWeights SQLite table (columns: TargetField, FeatureName, Weight, RecordCount, ComputedAt). Loaded into _computedWeights As Dictionary(Of String, Dictionary(Of String, Double)) during SuggestionEngine.New(). Threshold per target field: ≥3 distinct values each appearing ≥2 times (AnhaengeAblegen: both True and False ≥2 times). Fallback to per-target hardcoded weights (GetFeatureWeightsForXxx) when threshold not met or no stored weights exist. Recalculation is triggered every 50th saved session.

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
- Feature weights are defined per suggestion target in Core/SuggestionEngine.vb (separate weight profiles for ProjektPfad, ProjektstrukturPfad, Titel, AblageordnerSchema, MsgDateinameSchema, AnhaengeAblegen) and are applied independently during score aggregation.
- For Titel suggestions, Betreff dominates (0.55) so that semantic mail content drives the Titel choice. ProjektstrukturPfad weight is 0 to prevent a self-reinforcing cascade bias (ProjektstrukturPfad was itself suggested from the same historical record earlier in the cascade, so giving it weight here would circular-reinforce the same record). Ablageordner and Titel weights are 0 because both are mutable features initialized to zero and not yet set in the cascade when Titel is suggested.
- Distance lists are initialized with zeros at construction (length = historical record count). Fixed-feature subs overwrite them in PrepareSession(); mutable-feature subs overwrite them incrementally as the user edits fields.
- Current feature weighting is heuristic constants in Core/SuggestionEngine.vb and may need tuning with real usage data.
- Core/SuggestionEngine.vb relies on explicit framework imports (System, System.Collections.Generic, System.Linq) and uses a VB Char literal "\"c as one of the token separators in TokenizeForSimilarity.
- In Core/InputChecker.vb, ablageOrdnerPfad is Path.Combine(projektstrukturPfad, session.AblageordnerAufgeloest) where projektstrukturPfad is already the fully-combined absolute path. (A previous version incorrectly prepended projektPfad a second time, duplicating path segments — fixed.)
- Text encoding/comments show mixed umlaut encoding artifacts in several files; prefer preserving existing file encoding unless intentionally normalizing.
- SuggestionEngine.New() wraps GetAllSessionRecords() in try-catch and defaults to an empty list on failure so the Lazy never becomes permanently broken.
- EmbeddingService implements IDisposable; the InferenceSession is released via Dispose().
- SuggestionEngine implements IDisposable and exposes DisposeSharedInstance(); ThisAddIn_Shutdown calls it so ONNX resources are deterministically released when Outlook exits.
- GetEmbeddingService() returns Nothing if EmbeddingService construction fails (model files missing); all callers use null-conditional access so embedding gracefully degrades to zero similarity.
- MailUtils and InputChecker now release Outlook COM objects (Explorer, MailItem, Attachment) via Marshal.FinalReleaseComObject in Finally blocks to reduce long-running COM reference buildup.
- Dispatcher.BeginInvoke calls in MailDropWpfTaskPane guard against Dispatcher.HasShutdownStarted before posting.
- DirectoryTreeHelper.CreateDirectoryNodeWithExpand accepts an optional maxDepth parameter (default 20) to prevent StackOverflow from deep or cyclic directory structures.
- Session.HandleProjektSelection no longer takes a uiContext parameter (was unused).
- Session.CancelSession is implemented as Reset().
- AttachmentRenameDialog validates the filename for empty/invalid characters before accepting the dialog.
- ListBox1 in MailDropWpfTaskPane.xaml binds SelectedItem to ProjektPfad (was incorrectly bound to ProjektstrukturPfad, which caused a spurious cascade on every project selection).

## Conventions for Claude Code contributions

- Keep domain terms and UI labels in German to match existing code and user-facing text.
- Favor minimal, targeted edits; do not refactor broad areas unless needed for the requested change.
- Preserve the Session-based flow (PrepareSession/ProcessSession) and avoid introducing parallel state objects.
- For UI behavior changes, keep data-binding centered on Session properties and existing control names.
- For persistence changes, keep SessionRecord schema and SQLite migration impact explicit.
- When touching SuggestionEngine, keep behavior behind existing calls to avoid breaking current filing flow.
- Whenever code is changed, also review and update this CLAUDE.md so implementation notes, behavior descriptions, and caveats stay in sync.

## Response style and process transparency

- Answer critically, not just confirmatively. If a request is ambiguous, incomplete, or conflicts with existing architecture, say so explicitly before implementing.
- Every response that involves a task must state clearly whether: (a) a concept described in this CLAUDE.md or in the conversation was **translated into code** (no conceptual change), or (b) the concept itself was **updated** (the design changed during implementation). If (b), update the relevant section of CLAUDE.md to reflect the actual implemented design before closing the task.
- At the end of every response, include a short **"Next steps"** suggestion — one to three concrete items the project should address next, whether that means writing more code, refining a concept, or updating documentation. Base it on the current state of the codebase and known open issues.

## Quick verification checklist after changes

- Build succeeds in Visual Studio.
- Ribbon button opens the task pane.
- Exactly one selected mail enables editing; other selections disable editing.
- Selecting ProjektPfad refreshes Projektstruktur TreeView.
- Context menu actions (`Neuer Ordner`, `Loeschen`, `Umbenennen`) work on the expected TreeView node and keep selection in sync.
- `Neuen Ordner` button creates a folder in the expected TreeView location and auto-selects it.
- Placeholder fields resolve correctly on focus loss.
- OK saves .msg to expected folder.
- Optional attachment save works and handles long names.
- Session row is written to SQLite database.



