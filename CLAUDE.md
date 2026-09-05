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
- ProjektstrukturPfad TreeView supports creating, deleting, and renaming folders via context menu actions (Neuer Ordner, Loeschen, Umbenennen), and also supports creating folders via a Neuen Ordner button; if a node is selected, operations target that node, otherwise creation defaults to ProjektPfad.
- Creating, deleting, and renaming a folder (Neuer Ordner/Neuen Ordner, Loeschen, Umbenennen) no longer trigger a full BuildDirectoryTree() rebuild of the whole TreeViewData tree. Instead, all three actions call the same MailDropWpfTaskPane.RefreshTreeChildren(parentPath) helper right after the filesystem change, which calls Session.RefreshTreeViewChildren -> DirectoryTreeHelper.RefreshChildren: this re-reads only the direct children of the affected parent folder from disk (Directory.GetDirectories) and replaces just that node's Children collection (or the tree's root-level collection, if the parent is ProjektPfad itself) - reusing CreateDirectoryNodeWithExpand per rediscovered child so nested subtrees are populated normally. Everything outside that one parent's children (siblings, ancestors, other branches) is left untouched and keeps its expanded/collapsed TreeViewItem state. This one mechanism intentionally covers all three actions uniformly - the caller only needs to know which folder's contents changed (the parent of the created/deleted/renamed node), not how the tree needs to change, which is more robust than action-specific node surgery (in particular for rename, where a manual approach would otherwise have to update Name/FullPath/RelativePath on the renamed node and recursively on every descendant). The trade-off: nodes strictly below the affected parent are fully rebuilt and lose their own expand/collapse state (IsExpanded is reset to the level<=2-relative-to-that-parent default via CreateDirectoryNodeWithExpand) - this is scoped to the folder actually being edited, not the rest of the tree. If the affected parent is ProjektPfad itself, this means the entire root level (and everything below it) is rebuilt, since "the parent's children" is the whole first level in that case. RefreshTreeViewChildren/RefreshChildren return False if the parent folder can't be located in the current tree (should not normally happen); the UI helper falls back to the old full BuildDirectoryTree() rebuild in that case.
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
- SuggestionEngine exposes SuggestProjektPfad, SuggestProjektstrukturPfad, SuggestTitel, and SuggestAblageordnerSchema, all backed by shared helpers. SuggestAblageordnerSchema and SuggestMsgDateinameSchema fall back to the DefaultSchemaTemplate (`[Datum (formatiert)]_[Absender (kurz)]_[Titel]`) when no suggestion passes the threshold, so schema fields always get a sensible default even on first use.
- SuggestProjektPfad and SuggestProjektstrukturPfad iterate candidates in score order (via FindRecordsSortedByScore) and return the first whose path actually exists on disk, skipping non-existent paths while still respecting the SuggestionScoreThreshold. This avoids stale historical paths breaking the cascade.
- Internally, SuggestionEngine keeps FindBestRecordByField (returns single best record) for all fields except ProjektPfad/ProjektstrukturPfad, and FindRecordsSortedByScore (returns all above-threshold records sorted descending) for the two path fields. Avoids Func-signature overloading to prevent ambiguous lambda resolution in VB.
- For tracing/logging, SuggestionEngine passes explicit suggestion labels (e.g., ProjektPfad, MsgDateinameSchema) into the shared FindBestRecord helper instead of relying on lambda Method.Name.
- Session exposes SuggestXxx(value) methods that set the field value and then set the corresponding IsSuggestedXxx boolean flag to True. Cascade is driven by property setters: ProjektPfad → ProjektstrukturPfad → Titel → AbsenderKurz → AblageordnerSchema → MsgDateinameSchema → AnhaengeAblegen. Session also exposes IsSuggestedProjektPfad, IsSuggestedProjektstrukturPfad, IsSuggestedTitel, IsSuggestedAbsenderKurz, IsSuggestedAblageordnerSchema, IsSuggestedMsgDateinameSchema, IsSuggestedAnhaengeAblegen as full properties with PropertyChanged notification.
- MailDropWpfTaskPane.Session_PropertyChanged reacts to IsSuggestedXxx becoming True by showing the ✦ sparkle next to the field and (for ProjektstrukturPfad) dispatching a programmatic TreeView node selection via Dispatcher.BeginInvoke. The sparkle fades out when the user focuses/interacts with the field, which also sets the corresponding IsSuggestedXxx flag back to False (including AbsenderKurz, msg Dateiname and Anhänge ablegen).
- Default value for AnhaengeAblegen is True (checkbox always pre-checked). If the selected mail has no attachments, the checkbox is forced to False and grayed out (non-editable). Session.HasAnhaenge (read-only, computed from Anhaenge.Count > 0) drives this; SetEditMode special-cases CheckBoxAnhaenge to isEditable AndAlso Session.HasAnhaenge. The cascade's AnhaengeAblegen step is also skipped when HasAnhaenge is False.
- The attachment list (AttachmentScrollViewer) is controlled from code-behind: SetEditMode sets it to isEditable AndAlso Session.AnhaengeAblegen; Session_PropertyChanged(AnhaengeAblegen) updates it live so the list greys out immediately when the user unchecks the checkbox (even when attachments exist). The IsEnabled binding was removed from XAML to prevent SetEditMode from overriding it.
- ProjektPfad list shows the last 10 unique project paths (from DB) plus "anderes..." at the bottom, with a scrollable ListBox (MaxHeight="220"). DatabaseUtils.GetLastProjektVerzeichnisseForUser now exits the reader loop at count 10 (was 4).
- A private _applyingTreeViewSuggestion flag guards TreeView1_SelectedItemChanged against treating the engine's programmatic node selection as a user interaction (which would reset IsSuggestedProjektstrukturPfad and hide the sparkle immediately).
- The info popup (help button `i`) is organized into four headed sections: "Ablageordner und msg Dateiname" (schema/placeholder usage, including the quoting syntax for fixed text (`"..."`) and how an unquoted separator between two placeholders is dropped when a neighboring placeholder is empty - see Placeholder reference below), "Vorschläge" (suggestions are generated for all fields and marked with ✦, based on prior mail filings), "Vorschlagsdaten exportieren" (SuggestionRecord transfer, copy `%APPDATA%/MailDrop/sessions.db`), and "Vorschläge verbessern" (contains the "Gewichte neu berechnen" button with explanation; the button is disabled while recalculation runs and re-enabled on completion via Dispatcher.Invoke).
- InfoPopup.xaml's visual style (indigo `#4F46E5` section-header accents, `#374151`/`#6B7280` body/muted grays, `#D1D5DB`/`#E5E7EB` borders, 4px CornerRadius, flat `SecondaryButton` style) is duplicated from MailDropWpfTaskPane.xaml's `UserControl.Resources` rather than shared, since the popup is a separate top-level `Window` and the project has no `App.xaml`/merged resource dictionary to share styles from. Placeholder variables render as rounded "chip" Borders in a WrapPanel instead of a plain vertical list. If the task pane's palette or button styles change, InfoPopup.xaml's copies need updating to match by hand.
- TextBoxAblageordner and TextBoxMsgDateiname each carry a short static WPF ToolTip (hover) summarizing the schema syntax in generic terms (`[Variable]` wird durch den Wert ersetzt; `"Text"` in Anfuehrungszeichen bleibt immer erhalten; ein Trennzeichen ohne Anfuehrungszeichen zwischen zwei `[Variable]` entfaellt, wenn eines leer ist) plus a pointer to the info popup for the full explanation and examples. Deliberately kept to 3-4 lines rather than duplicating the full info-popup text, since a hover tooltip disappears on focus/interaction and isn't meant for reading at length.
- Task pane button order (left to right): OK → Abbrechen → i.

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
- Error log (Release-safe): %APPDATA%/MailDrop/error.log, written by Helpers/Logger.vb's LogError(context, ex). Release builds compile out Debug.WriteLine entirely (DefineDebug=false in MailDrop.vbproj), so this file is currently the only diagnostic trail available on an end-user machine; check it first when "the MailDrop button does nothing" is reported. Logger.LogError never throws itself (wrapped in Try/Catch) so logging failures cannot mask the original error or crash the caller.
- SQLite schema versioning uses `PRAGMA user_version`; `SessionDatabaseManager.CurrentSchemaVersion` defines the expected schema and `ApplyMigrations(...)` upgrades older databases.
- Database manager initialization is lazy: ThisAddIn.CurrentDatabaseManager creates SessionDatabaseManager on first access (not in ThisAddIn_Startup).
- ONNX model files are loaded from output folder path Models/model.onnx and Models/vocab.txt.
- The root-level model.onnx and vocab.txt are source artifacts; runtime inference uses the files under Models/.

## Distribution (ClickOnce / GitHub Releases zip + network drive)

- Distribution channel: the ClickOnce `Publish` output (`setup.exe`, `MailDrop.vsto`, `Application Files/`, `Install-Certificate.ps1`) is zipped and attached to a GitHub Release; the same extracted zip is also copied onto a network drive. Anyone who wants to install runs `setup.exe` from either location. The `Publish/` folder itself is **no longer committed to git** (build output, `.gitignore`d) except `Publish/Install-Certificate.ps1`, which is hand-written source and stays tracked.
- Auto-update: `UpdateEnabled=true` and `UpdateMode=Background` in `MailDrop.vbproj`, with no hard-coded `UpdateUrl`/`InstallUrl`. ClickOnce then defaults to using the location an install ran *from* as its update source. Installing directly from the network drive means that path becomes the client's update source (checked in the background on every Outlook start, applied on next restart). Installing from a locally extracted GitHub Releases zip gets no auto-update; updating means reinstalling from a newer zip. `IsWebBootstrapper=False` since the network-drive/offline install still shouldn't assume internet access for .NET/VSTO prerequisites.
- `PublisherName` is set to `MailDrop` (affects the text shown in the ClickOnce trust/install dialog).
- VSTO add-ins require signed ClickOnce manifests: MSBuild's `Microsoft.VisualStudio.Tools.Office.targets` hard-fails (`Cannot build because the ClickOnce manifest signing option is not selected`) if `SignManifests` is false, unlike plain ClickOnce (non-Office) projects where unsigned manifests are allowed. `SignManifests=false` was tried and confirmed to break the Release publish build; it is not a viable option for this project.
- MailDrop.vsto/.dll.manifest are therefore signed with a self-created, self-signed certificate (`MailDrop_1_TemporaryKey.pfx`, subject `CN=MailDrop`, deliberately generated with a 30-year validity via `New-SelfSignedCertificate -NotAfter (Get-Date).AddYears(30)` rather than VS's default 1-year "temporary key", so client trust never needs to be redone due to expiry). This certificate is not issued by a trusted CA, so installation on any machine other than the signing machine fails with a certificate/publisher-verification error.
- `Publish/Install-Certificate.ps1` fixes this: it embeds the public certificate (extracted right after certificate generation, no private key involved) and imports it into the current user's `Root` and `TrustedPublisher` certificate stores. This uses the `CurrentUser` store scope, which requires **no administrator rights** (chain building consults both `CurrentUser` and `LocalMachine` Root stores, so `CurrentUser\Root` alone is sufficient for that user's trust decisions). Supports `-Uninstall` to remove it again. End users run this once before running `setup.exe`; the script ships alongside `setup.exe` in the zip/network-drive folder.
- The embedded certificate expires 2056-07-01. It only needs regenerating if the private key is ever compromised or replaced; in that case `Install-Certificate.ps1`'s embedded `$certBase64` block, `MailDrop_1_TemporaryKey.pfx`, and the `ManifestCertificateThumbprint` in `MailDrop.vbproj` must all be regenerated together, and every already-installed user must re-run the script for the new thumbprint (this is the one structural downside of self-signing vs. a CA-issued cert: cert rotation is not transparent to already-trusted clients).
- README.md / README.en.md document the end-user install flow (run `Install-Certificate.ps1`, then `setup.exe`) and the auto-update behavior above.
- No GitHub Actions workflow exists yet for building the Release build, zipping `Publish/`, or creating the GitHub Release — this is currently a manual process (publish in Visual Studio, zip the `Publish/` folder, `gh release create` or the GitHub UI, then manually copy the extracted zip to the network drive).
- `SQLite.Interop.dll` (native, from the `Stub.System.Data.SQLite.Core.NetFramework` NuGet package) is physically committed at `x86\SQLite.Interop.dll` and `x64\SQLite.Interop.dll` in the project root (copied from `packages\Stub.System.Data.SQLite.Core.NetFramework.<version>\build\net46\x86\`/`x64\`), referenced via plain `<Content Include="x86\SQLite.Interop.dll">`/`x64\...` items (`CopyToOutputDirectory=PreserveNewest`, no `<Link>`) in `MailDrop.vbproj` — same pattern as `Models\model.onnx`. `.gitignore`'s stock VS template ignores `x86/`/`x64/` build-output folders, so it carries explicit `!/x86/`, `!/x86/SQLite.Interop.dll`, `!/x64/`, `!/x64/SQLite.Interop.dll` exceptions to keep these two files tracked. A `<Link>`-based approach (redirecting `<Content Include="packages\...\SQLite.Interop.dll">` to `x86\SQLite.Interop.dll` via `<Link>`) was tried first and technically worked when publishing via raw `msbuild /t:Publish`, but Visual Studio's own Publish command rejects `<Link>` for any source file already inside the project directory tree (`packages\` is inside the project root) with a "kann kein Link hinzugefügt werden, da sie sich im Projektverzeichnisbaum befindet" warning and skips the file — hence the physical-copy approach instead. Without either fix, ClickOnce's publish/manifest step (which only tracks explicit project `<Content>`/`<None>` items, not files produced by an imported package `.targets` file) deploys the interop DLL at its nested source-relative path (`packages\Stub.System.Data.SQLite.Core.NetFramework.<version>\build\net46\x86\SQLite.Interop.dll`) instead of the flattened `x86\`/`x64\` path that `System.Data.SQLite`'s native-library probing requires — the DLL is present in the installed ClickOnce cache but at the wrong relative location, causing a `DllNotFoundException` on `SessionDatabaseManager.EnsureDatabaseExists()` that only reproduces from an actual ClickOnce-installed copy (a direct `bin\Release` load works regardless, since the package's own `.targets` import additionally copies a correctly-flattened copy there that ClickOnce never sees). If the NuGet package version changes, re-copy both files from the new package version's `build\net46\x86\`/`x64\` folder.
- **Publishing from the Visual Studio IDE (File > Publish, not raw `msbuild /t:Publish`) can silently revert the two `SQLite.Interop.dll` `Content` items back to the broken `<Link>`-to-`packages\...` form**, re-triggering both the "kann kein Link hinzugefügt werden" warning and the underlying `DllNotFoundException` risk on the next install. Observed once already: after committing the physical-copy fix above, a single VS-driven Publish (which also auto-bumps `ApplicationRevision`) reintroduced the exact pre-fix `<Content Include="packages\...\build\net46\x64\SQLite.Interop.dll"><Link>x64\SQLite.Interop.dll</Link>...` lines verbatim, independent of anything committed. Root cause not fully pinned down, but a defensive `<ItemGroup><SQLiteInteropFiles Remove="@(SQLiteInteropFiles)" /></ItemGroup>` right after the package's `Stub.System.Data.SQLite.Core.NetFramework.targets` import (plus a `CollectSQLiteInteropFiles=false` property) is in place to stop the package's own `CollectSQLiteInteropFiles` target from re-feeding these files into VSTO's `CopyPublishItems` pipeline — neither alone fixed the observed recurrence (the literal, persisted `Content` lines were the actual cause that one time), so **always diff `MailDrop.vbproj` after any VS-driven Publish** and re-apply the plain `x86\`/`x64\SQLite.Interop.dll` `Content` items (no `<Link>`) if the `packages\...`/`<Link>` form has reappeared, before committing.
- Any change to `MailDrop.vbproj`'s `ManifestCertificateThumbprint`/`ManifestKeyFile` or `PublisherName` changes the ClickOnce deployment identity. Re-registering/reinstalling afterward on a machine that already has an older-identity install fails with `Microsoft.VisualStudio.Tools.Applications.Deployment.AddInAlreadyInstalledException` ("eine andere Version installiert... zuerst uber Software deinstallieren"), even for a same-location, higher-version install — the stale registration(s) must be removed first via `VSTOInstaller.exe /Uninstall <original-install-url>` (find stale entries via `HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*` filtered to DisplayName `MailDrop`; get each one's original install URL from its `UninstallString`) before a fresh `/Install` of the new identity will register.
- Outlook auto-disables an add-in after repeated load/runtime failures, recorded per-user at `HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems` (one value per disabled item, name decodes from UTF-16 to the add-in's DLL path). `DoNotDisableAddinList` (also under `...\Resiliency\`) only prevents *future* auto-disabling; it does not undo an existing `DisabledItems` entry, which must be deleted manually (then restart Outlook) to re-enable.

## Branching model

- `released`: stable branch intended for production-ready releases.
- `development`: integration branch for ongoing development changes. Routine changes are committed directly to `development` (no per-change feature branch/PR required); an isolated `feature/...`/`fix/...` branch is only used when a change is large/risky enough to want it reviewed in isolation before landing.
- Release-ready states on `development` are promoted into `released` via pull request (this is the one gate that stays PR-only, since `released` is what end users install from).
- Repository governance details (review gates and merge flow) are documented in CONTRIBUTING.md.
- GitHub default branch should be `released` (currently `development` is the default on the remote — switch it in repository settings).

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
- Feature weights are data-driven per target field when sufficient history exists. Weights are computed via pairwise Pearson correlation between feature similarity vectors and a binary label (same/different target value for each record pair), clipped to [0, ∞) and normalized to sum 1. Stored per target field in the ComputedWeights SQLite table (columns: TargetField, FeatureName, Weight, RecordCount, ComputedAt). Loaded into _computedWeights As Dictionary(Of String, Dictionary(Of String, Double)) during SuggestionEngine.New(). Threshold per target field: ≥3 distinct values each appearing ≥2 times (AnhaengeAblegen: both True and False ≥2 times). Fallback to per-target hardcoded weights (GetFeatureWeightsForXxx) when threshold not met or no stored weights exist. Recalculation is triggered every 50th saved session and also via the "Gewichte neu berechnen" button in the task pane.
- Cascade-aware feature sets: for each target field, only features causally available at that cascade stage are included in the correlation computation. Fixed input features (Betreff, Datum, AbsenderDomain, Absender, AusfueBenutzer, AusfueDatum) are always available. Mutable features become available in cascade order: ProjektPfad from ProjektstrukturPfad onwards, ProjektstrukturPfad from Titel onwards, Titel from AbsenderKurz onwards, Ablageordner from MsgDateinameSchema onwards. The mapping is encoded in MutableFeatureAvailableFromStep (Dictionary keyed by feature name, value = CascadeStep enum) — the cascade order itself is not redefined, only the input-availability dependency. PearsonCorrelation returns Double.NaN when the feature vector is constant (all pairs have identical similarity), handled in debug output as "Feature-Vektor konstant". RecalculateWeightsFromHistory emits structured Debug.WriteLine output per target field: K/M threshold info, per-feature Pearson value with reason, skipped cascade features, and normalized before/after weight comparison.

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
- Fixed/literal text in a schema can be forced by wrapping it in double quotes, e.g. [Titel]" (Entwurf) "[Absender (kurz)] - the quoted text is always emitted verbatim (quotes stripped), regardless of whether neighboring placeholders are empty. Any other (unquoted) text that sits directly between two placeholders - e.g. the "_" in the default schema - is treated as a connector/separator rather than fixed text, and is only emitted when it actually ends up joining two placeholders that both resolved to a non-empty value (see below). Unquoted text that does not sit between two placeholders (e.g. a fixed prefix before the first placeholder) is not a connector and is always emitted unchanged, quotes or not.
- Session.ReplacePlaceholders (Core/Session.vb) parses the schema template into SchemaSegment values of kind Literal (quoted text, or any text not sandwiched between two placeholders), Connector (unquoted text directly between two placeholders), or Placeholder (a resolved value, possibly empty), via a single regex matching both a quoted-string pattern and the known placeholder tokens; unmatched substrings between tokens ("gaps") are classified as Connector only when both their immediate neighbors are Placeholder segments, Literal otherwise. Resolution is then a "skip-empty-join": a Connector is held back (pendingConnector) instead of appended immediately, and only actually emitted right before the next non-empty Placeholder that follows a previously-emitted non-empty Placeholder; an empty Placeholder is skipped without discarding a still-pending connector (a later Connector simply overwrites it). This guarantees exactly one connector between any two placeholders that actually resolved to a value, however many empty placeholders sit between them, and never leaves a leading or trailing connector - independent of quoting. Example: "[A]_[B]_[C]" with only B empty resolves to "A_C" (not "AC" or "A__C"); with both B and C empty it resolves to "A" (no trailing "_"); [A]" - "[B] with B empty still resolves to "A - " because the quoted " - " is fixed text, not a connector.

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
- ThisAddIn.MailAblegen_Click, the Explorer.SelectionChange handler, and PreloadTaskPaneInBackground each wrap their work in Try/Catch and call Logger.LogError on failure, instead of letting exceptions propagate to the VSTO ribbon/event boundary. This matters because Outlook/VSTO silently swallows unhandled exceptions thrown from ribbon callbacks and COM event handlers (no dialog, no visible effect) unless the user has enabled Outlook's "Show add-in user interface errors" option — without the Try/Catch, a failure anywhere in MailSelected() -> Session.PrepareSession() (e.g. first-time SQLite/ThisAddIn.CurrentDatabaseManager init failing) meant the button click did nothing observable and left no trace. taskPane.Visible = True in MailAblegen_Click now always runs as long as the task pane itself was created, even if populating it (MailSelected) failed, so the pane is shown (possibly not fully populated) rather than the click appearing to do nothing; if task pane creation itself fails, a MessageBox is shown pointing at error.log since there is nothing else to display.
- MailDropWpfTaskPane.SetEditMode(isEditable) iterates all controls via FindVisualChildren(Of Control) and now wraps each control's IsEnabled assignment in its own Try/Catch (logs via Logger.LogError and continues) instead of one Try/Catch around the whole loop, so one control failing to update can't leave later controls stuck at their previous IsEnabled value. FindVisualChildren itself also wraps each VisualTreeHelper.GetChildrenCount/GetChild call individually (Yield is not allowed inside a VB Try block that has a Catch, so the recursive descent and the error handling had to be split into separate Try blocks with no Yield inside them): a subtree that fails to enumerate (e.g. ListBox1/TreeView1 item containers not yet generated) is logged and skipped instead of aborting the entire enumeration, which previously could leave arbitrary controls after the failure point permanently disabled/greyed out with no trace of why.
- DirectoryTreeHelper.CreateDirectoryNodeWithExpand accepts an optional maxDepth parameter (default 20) to prevent StackOverflow from deep or cyclic directory structures.
- Session.HandleProjektSelection no longer takes a uiContext parameter (was unused).
- Session.CancelSession is implemented as Reset().
- AttachmentRenameDialog validates the filename for empty/invalid characters before accepting the dialog.
- ListBox1 in MailDropWpfTaskPane.xaml binds SelectedItem to ProjektPfad (was incorrectly bound to ProjektstrukturPfad, which caused a spurious cascade on every project selection).
- DuplicateWarningNotification, OverwriteWarningNotification, and ErrorNotification are deliberately all placed in the same Grid.Row as SuccessNotification (stacked in one cell, shown one at a time, Opacity animated 0→1). If a row is ever inserted into the outer Grid.RowDefinitions above these, every one of these Grid.Row values (not just SuccessNotification's) must be bumped together with the OK/Abbrechen/i StackPanel's row - a mismatch previously left the notification Borders sitting on top of (later in Z-order than) the buttons in the same row, and since Opacity="0" does not disable hit-testing in WPF, the invisible Borders silently swallowed all clicks on OK/Abbrechen/i even though nothing was visibly wrong.
- SuggestProjektPfad and SuggestProjektstrukturPfad use FindRecordsSortedByScore to iterate candidates above SuggestionScoreThreshold in score order, skipping records whose paths don't exist on disk. The first existing path wins; if none qualifies, an empty string is returned (no suggestion).
- AnhaengeAblegen default is True. PrepareSession forces it to False (and clears IsSuggestedAnhaengeAblegen) when the mail has no attachments, as a safety net after the cascade. Session.HasAnhaenge (Anhaenge.Count > 0) is raised via OnPropertyChanged after ReadAttachmentNames and after Anhaenge.Clear() in Reset.

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



