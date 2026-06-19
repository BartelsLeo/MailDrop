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

- productive: Stable branch for production-ready releases
- development: Integration branch for ongoing development
- feature/*: Short-lived working branches merged into development
- Release flow: Mature states from development are promoted to productive

## Repository Governance (Best Practice)

- Default branch in GitHub: productive
- Pull request targets:
  - Normal flow: feature/* -> development
  - Release promotion: development -> productive
  - Hotfix flow: hotfix/* -> productive, then back-merge into development
- Protection rules for productive:
  - No direct pushes
  - Pull request merge only
  - At least 1 required review
  - Dismiss stale reviews when new commits are pushed
- Protection rules for development:
  - No direct pushes
  - Pull request merge only
  - At least 1 review recommended

Note: The actual default-branch switch and branch protection are configured in GitHub repository settings.

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
