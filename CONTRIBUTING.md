# Contributing to MailDrop

This project follows a two-branch delivery model for safe releases.

## Branch model

- productive: Production-ready branch
- development: Integration branch for daily work
- feature/*: New features and refactors
- fix/*: Non-critical bug fixes
- hotfix/*: Urgent production fixes

## Pull request flow

1. Create your branch from development (or from productive for hotfixes).
2. Keep changes focused and small.
3. Open a pull request:
   - feature/* or fix/* -> development
   - hotfix/* -> productive
4. After hotfix merge into productive, create a follow-up PR from productive (or hotfix branch) back to development.
5. For release promotion, open PR from development -> productive.

## Review and quality gates

- No direct pushes to productive or development
- Merge via pull request only
- At least one review before merge
- Resolve discussions before merge
- Prefer passing build/test checks before merge

## Commit guidance

- Use clear, action-oriented messages
- Keep unrelated changes out of the same PR
- Update documentation when behavior changes

## Release guideline

1. Stabilize in development.
2. Open PR development -> productive.
3. Tag release on productive (for example v1.2.0).
4. If needed, cherry-pick or hotfix from productive.

## Repository settings (one-time)

Configure these in GitHub repository settings:

- Set default branch to productive
- Protect productive:
  - Require pull request before merge
  - Require at least 1 approval
  - Dismiss stale approvals when new commits are pushed
  - Restrict direct pushes
- Protect development:
  - Require pull request before merge
  - Restrict direct pushes

