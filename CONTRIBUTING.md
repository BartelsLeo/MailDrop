# Contributing to MailDrop

This project uses a single-branch delivery model.

## Branch model

- `main`: the only ongoing branch — routine changes are committed directly here
- `feature/*` / `fix/*`: optional, only for a change large or risky enough to want isolated review before it lands in `main`

There used to be a separate `released` branch that `development` was promoted into via pull request before a release. It was retired: it added a promotion step without enough release cadence to justify a second branch, and has been deleted. `development` was later renamed to `main` (2026-09-20) once it was the repo's only ongoing branch, so the name no longer implied an unreleased/integration-only status. Releases are cut directly from `main` (see Release flow below).

## Day-to-day workflow

1. Commit directly to `main` for routine changes (bug fixes, small features, docs).
2. Use a `feature/*` or `fix/*` branch + PR into `main` only when you want the change reviewed in isolation first — this includes urgent fixes, which no longer need a separate `hotfix/*` flow now that there's no second branch to hotfix against.
3. Keep changes focused and small regardless of which path you use.

## Release flow

1. Stabilize changes on `main`.
2. Tag the release commit on `main` directly, matching the ClickOnce `ApplicationVersion` (for example `v1.0.1.0`).
3. Publish (Visual Studio ClickOnce Publish), zip `Publish/`, attach to a GitHub Release, and copy to the network drive — see the Distribution section in `CLAUDE.md`.

## Review and quality gates

- `main` allows direct pushes — no PR required for routine work.
- Prefer passing build/test checks before tagging a release on `main`.

## Commit guidance

- Use clear, action-oriented messages.
- Keep unrelated changes out of the same commit/PR.
- Update documentation when behavior changes.

## Repository settings

- Default branch: `main` (no separate protected branch to configure now that `released` is gone).
- `main` is left unprotected so routine work can be committed directly.
