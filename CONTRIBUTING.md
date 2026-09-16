# Contributing to MailDrop

This project follows a two-branch delivery model for safe releases.

## Branch model

- `released`: Production-ready branch
- `development`: Integration branch for daily work — routine changes are committed directly here
- `feature/*` / `fix/*`: Optional, only for a change large or risky enough to want isolated review before it lands in `development`
- `hotfix/*`: Urgent production fixes, branched from `released`

## Day-to-day workflow

1. Commit directly to `development` for routine changes (bug fixes, small features, docs).
2. Use a `feature/*` or `fix/*` branch + PR into `development` only when you want the change reviewed in isolation first.
3. Keep changes focused and small regardless of which path you use.

## Release flow

1. Stabilize changes on `development`.
2. Open a pull request `development` -> `released`.
3. Tag the release on `released`, matching the ClickOnce `ApplicationVersion` (for example `v1.0.1.0`).
4. For an urgent production fix: branch `hotfix/*` from `released`, PR it into `released` directly, then open a follow-up PR from `released` (or the hotfix branch) back into `development` so the fix isn't lost on the next promotion.

## Review and quality gates

- `released` is protected: merge via pull request only, no direct pushes, at least one review (or explicit self-review sign-off if solo) before merge.
- `development` allows direct pushes — no PR required for routine work.
- Prefer passing build/test checks before promoting `development` -> `released`.

## Commit guidance

- Use clear, action-oriented messages.
- Keep unrelated changes out of the same commit/PR.
- Update documentation when behavior changes.

## Repository settings (one-time)

Configure these in GitHub repository settings:

- Set default branch to `released`
- Protect `released`:
  - Require pull request before merge
  - Require at least 1 approval
  - Dismiss stale approvals when new commits are pushed
  - Restrict direct pushes
- `development` is left unprotected so routine work can be committed directly.
