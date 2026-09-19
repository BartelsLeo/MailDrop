# MailDrop

MailDrop is a VSTO add-in for Microsoft Outlook. It helps users file emails and optional attachments into project folder structures in a consistent way.

## Value

- Faster and more consistent email filing into project folders
- Fewer manual mistakes through placeholder resolution and input validation
- History-based suggestions (SuggestionRecords stored in SQLite)
- Repeatable workflow for teams with recurring projects

## Features

- Outlook ribbon button to open a right-side task pane
- Editing is enabled only when exactly one mail is selected
- Dynamic project structure tree with:
  - New folder
  - Delete (empty folders only)
  - Rename
- Placeholder workflow for:
  - Filing folder
  - msg filename
- Optional filing of all attachments
- SuggestionEngine with ONNX embeddings for cascading suggestions:
  - ProjektPfad (project root path)
  - ProjektstrukturPfad (relative project structure path)
  - Titel (title)
  - Absender (kurz) (short sender)
  - Ablageordner schema
  - msg filename schema
  - Store attachments (boolean)
- Sparkle indicators for automatically applied suggestions
- Help popup with:
  - Placeholder reference
  - Guidance for transferring SuggestionRecords via SQLite file

## Tech Stack

- Language: Visual Basic .NET
- Framework: .NET Framework 4.7.2
- Add-in technology: VSTO 4.0
- Host: Microsoft Outlook (desktop)
- Persistence: SQLite (System.Data.SQLite)
- ML/Embedding: ONNX Runtime + Tokenizer
- Build: Visual Studio 2022 / MSBuild

## Project Structure (Overview)

- Core/: Session logic, validation, SuggestionEngine
- Helpers/: DB access, tree helpers, mail helpers
- Services/: EmbeddingService
- UI/: Ribbon, task pane, dialogs, help popup
- Models/: model.onnx, vocab.txt (runtime model artifacts)

## Branching Strategy

- development: the only ongoing branch, direct commits for routine changes
- feature/*: short-lived working branches for changes that want isolated review, merged into development
- A previous separate `released` branch (promotion gate `development` -> `released` via pull request) was removed since it added a promotion step without enough release cadence to justify it; releases are now tagged directly on `development`

## Repository Governance (Best Practice)

- Default branch in GitHub: development
- Pull request target: feature/* -> development (only when isolated review is wanted, otherwise commit directly)
- Protection rules for development:
  - Direct pushes allowed for routine work
  - No pull request required for routine changes

See CONTRIBUTING.md for the full workflow, including release tagging.

## Installation and Setup

### Requirements

- Windows with Microsoft Outlook desktop installed
- VSTO Runtime
- Visual Studio 2022
- .NET Framework 4.7.2 Targeting Pack

### Run locally

1. Open MailDrop.sln in Visual Studio 2022.
2. Restore NuGet packages (packages.config based project).
3. Select build configuration Debug | Any CPU.
4. Start debugging (F5).
5. Visual Studio starts Outlook as host process and loads the add-in.

### ClickOnce installation (end users)

MailDrop is distributed as a downloadable zip via GitHub Releases; the same extracted zip
(`setup.exe`, `MailDrop.vsto`, `Application Files/`) is also placed on a network drive. Anyone who
wants to install MailDrop can run `setup.exe` either from the locally extracted zip or directly from
the network drive.

VSTO add-ins require signed ClickOnce manifests (MSBuild refuses to build otherwise, with
`Cannot build because the ClickOnce manifest signing option is not selected`); MailDrop is therefore
signed with a self-signed certificate (not one from a public certificate authority). On the
development machine that certificate is already trusted, but on any other machine the installation
fails or aborts with a certificate warning.

Fix: run `Install-Certificate.ps1` before installing (it ships next to `setup.exe` in the zip / on
the network drive). The script trusts the (public) MailDrop certificate for the current user; it
needs no private key and contains none, and requires **no administrator rights** (it only touches
the current user's certificate stores):

```powershell
powershell -ExecutionPolicy Bypass -File .\Install-Certificate.ps1
```

Then run `setup.exe`. The certificate is deliberately valid for 30 years (until 2056-07-01) so this
trust step never has to be repeated for already-installed users. Only if the certificate is ever
regenerated (e.g. private key compromise) does the script need to be regenerated from `MailDrop.vsto`
(see the comment in the script) and re-run by every user.

#### Auto-update

ClickOnce is configured with `UpdateEnabled=true` and no hard-coded update address. As a result, the
location an install ran from automatically becomes its update source:

- **Installed directly from the network drive**: MailDrop then checks that network drive in the
  background on every Outlook start (`UpdateMode=Background`) and picks up a newer version on the
  next restart. No extra step needed.
- **Installed from a locally extracted zip** (e.g. downloaded from GitHub Releases): there is no
  automatic update check. Updating requires manually installing again from a newer zip.

**Note:** Installing from a OneDrive-synced folder (e.g. `...\OneDrive - <Org>\...\MailDrop\`) is not a
tested/supported third option alongside the network drive/zip. Whether that's actually a contributing
cause of the failures described below is unconfirmed (see Troubleshooting) — avoided as a precaution
regardless.

## Usage

1. Select exactly one email in Outlook.
2. Click the MailDrop ribbon button.
3. Choose a project root folder or use anderes....
4. Select/create the project structure path in the tree.
5. Review fields (Titel, Absender kurz, Ablageordner, msg Dateiname).
6. Optionally enable attachment filing.
7. Click OK.

## Placeholders

Supported placeholders in Ablageordner and msg Dateiname:

- [Titel]
- [Absender]
- [Absender-Domain]
- [Empfaenger]
- [Empfaenger (kurz)]
- [Betreff]
- [Datum]
- [Datum (formatiert)]
- [Absender (kurz)]

## Data and SuggestionRecords

- SQLite database file: %APPDATA%/MailDrop/sessions.db
- This file stores session history and suggestion-relevant records.

### Transfer SuggestionRecords to another PC

1. Close Outlook (and MailDrop) on the source PC.
2. Copy %APPDATA%/MailDrop/sessions.db.
3. Close Outlook (and MailDrop) on the target PC.
4. Place the file at exactly %APPDATA%/MailDrop/sessions.db and replace the existing file.
5. Restart Outlook.

## Implementation (Architecture and Flow)

### Startup and UI

- Entry point: ThisAddIn_Startup
- Ribbon action calls MailAblegen_Click
- Task pane hosts MailDropWpfTaskPane

### Session flow

- PrepareSession:
  - Reset
  - Read mail metadata
  - Load recent project paths
  - Acquire shared SuggestionEngine
  - Calculate initial feature distances
  - Suggest and optionally apply ProjektPfad
- ProcessSession:
  - Validate input
  - Create target folder
  - Save mail as .msg
  - Optionally save attachments
  - Persist SessionRecord to SQLite

### SuggestionEngine (simplified)

- Historical SessionRecords are loaded from SQLite.
- Scoring is based on weighted features, including:
  - Subject semantic similarity
  - Date
  - Sender/domain
  - User
  - Project path/project structure path
  - Title/filing folder
- Suggestions are applied in a cascade along the session workflow.

## Validation and Constraints

- ProjektPfad must exist.
- ProjektstrukturPfad must exist under ProjektPfad.
- Invalid file/path characters are rejected.
- Path/file length checks are applied.
- Long attachment names can be handled via rename dialog.

## Troubleshooting

- MailDrop button is missing in Outlook ribbon:
  - Restart Outlook completely.
  - Check COM Add-ins in Outlook and ensure MailDrop is enabled.
  - Start once from Visual Studio debug mode to refresh add-in load behavior.
- Task pane does not open or is empty:
  - Ensure exactly one mail is selected.
  - Restart Outlook and test again.
  - Verify Debug | Any CPU build configuration.
- Suggestions do not appear:
  - Historical records are required (sessions.db must not be empty).
  - Verify output model files exist: Models/model.onnx and Models/vocab.txt.
  - Restart Outlook to reinitialize lazy-loaded components.
- Saving mail/attachments fails:
  - Verify project root and project structure path exist.
  - Check for invalid path/file characters or path length limits.
  - Use attachment rename dialog for long filenames.
- SuggestionRecords did not transfer to another PC:
  - Ensure Outlook was closed on both PCs during file copy.
  - Use exact path: %APPDATA%/MailDrop/sessions.db.
  - Confirm existing target file was actually replaced.
- Update fails ("update not possible") even though a fresh install from the same location works,
  and/or the add-in gets auto-disabled right after install due to a timeout, possibly with this
  error: `DeploymentDownloadException` / `UnauthorizedAccessException: Der Zugriff auf den Pfad
  "...\AppData\Local\Temp\Deployment\...\MailDrop.dll" wurde verweigert`:
  - Confirmed cause (readable directly from the stack trace): ClickOnce first writes downloaded
    files into a randomly-named, purely local folder under `%LOCALAPPDATA%\Temp\Deployment\`, then
    reopens that file to verify the manifest/signature. An access error at exactly that moment is a
    well-known, generic ClickOnce failure mode, usually caused by antivirus/EDR real-time scanning
    briefly locking newly written DLL/EXE files (the deployment stack does not retry). This can
    happen regardless of the install source, since the failing file is local.
  - One observed case had an install source inside a **OneDrive-synced folder** (e.g.
    `...\OneDrive - <Org>\...\MailDrop\`), which is neither of the two tested/supported install
    locations (network drive, or a zip extracted elsewhere). Whether OneDrive as the source actually
    contributes to triggering the antivirus lock race is **not confirmed** — the failing file itself
    is local and outside any OneDrive folder — so this is more a risk factor / reason to stick to
    the supported locations than a proven root cause.
  - This happens inside Outlook's own ClickOnce/VSTO loader, before this project's own add-in code
    even starts — it is **not** fixable via a code change in this repository.
  - Fix, in order of evidence: add an antivirus/EDR exclusion for `%LOCALAPPDATA%\Temp\Deployment\`
    (the best-evidenced fix for this exact exception), clear the ClickOnce cache
    (`rundll32.exe dfshim.dll CleanOnlineAppCache`), remove the stale `DisabledItems` registry value
    for MailDrop under `HKCU:\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems`, then
    reinstall. Also install/update only from the network drive or a zip outside OneDrive as a
    precaution, since that combination is untested.

## Verification Checklist

- Build succeeds in Visual Studio
- Ribbon button opens task pane
- Editing enabled only for exactly one selected mail
- Project tree reacts correctly to selection and folder actions
- Placeholders resolve on focus loss
- OK saves .msg as expected
- Optional attachment save works
- Session is written to SQLite

## License

There is currently no explicit license file in this repository.
If needed, add a LICENSE file.

## Contributing

The canonical development and PR workflow is documented in CONTRIBUTING.md.
