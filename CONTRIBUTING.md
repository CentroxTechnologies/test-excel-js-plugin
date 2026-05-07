# Contributing to PowerPair

Thanks for picking this up. This is the contribution flow we follow on this repo. Read once, refer back when needed.

## Where to start

1. Read [`README.md`](./README.md) for what PowerPair is and how to install / run it.
2. Read [`DEV-GUIDE.md`](./DEV-GUIDE.md) for project layout, request flow, and how to add a new action type or LLM provider.
3. Pick an item from "What's next" in the README.

## Branching + commits

- Cut a feature branch off `main`: `git checkout -b feat/<short-name>` or `fix/<short-name>`.
- Keep commits focused. One logical change per commit when possible. Squash on merge if you have noisy WIP commits.
- Commit subject under 72 characters, imperative mood ("Add validator for sort_range", not "Added a thing").
- Body explains the *why*, not the *what*, the diff already shows the what.

## Pull requests

- Open a PR against `main`. Use the [PR template](./.github/PULL_REQUEST_TEMPLATE.md), it's pre-filled with the checklist.
- Link any related issue (`Closes #N`).
- Keep PRs small. Anything over ~400 lines of diff should probably be split.
- CI runs on every PR (Python syntax + JS syntax + manifest validation). Don't merge red.
- At least one approving review from a CODEOWNER before merging. Self-merging is fine for trivial doc fixes only.

## Code style

- **Python (backend)**: 4-space indent, type hints on public functions, `ruff format` if you have it. No mock-only code paths in production files, mocks live in tests.
- **JavaScript (addin)**: 2-space indent, semicolons, no TypeScript yet (intentional, keep the scaffold readable).
- **Comments**: explain *why*, not *what*. The code shows what.
- **Files over 300 lines**: consider splitting before merging.

## Adding a new feature

The README lists the next priorities. If you're adding something not on that list, open an issue first to discuss scope. Drift is the enemy.

Common feature types and their patterns:

- **New AI action type** (e.g. `apply_filter`): 4-step recipe in `DEV-GUIDE.md`, system prompt + validator + Excel handler + mock handler.
- **New LLM provider** (e.g. Gemini): subclass `AIProvider` in `backend/ai_engine.py`, plumb via `get_provider()`, add env vars to `.env.example`.
- **UI changes**: update `addin/src/taskpane/{html,css,js}` and redeploy the `gh-pages` branch (the live hosted plugin).

## Testing

There's no automated test suite yet, we rely on manual smoke checks per the table in `DEV-GUIDE.md`. If you're adding a feature that's easy to unit-test (validator rules, formula generation), please add a test file alongside your change.

## Reporting bugs

Use the [bug report template](./.github/ISSUE_TEMPLATE/bug_report.md). Include:
- What you typed in the sidebar
- What action_type came back
- What Excel did vs what you expected
- Browser + Excel surface (Online / Desktop Win / Desktop Mac)

## Keys + secrets

- Never commit `backend/.env` (it's gitignored, keep it that way).
- API keys live only in `.env` on the developer machine or in environment variables on the deploy target.
- Don't paste keys into PR descriptions, issues, or commit messages. If a key leaks, rotate it immediately.

## Questions

Ping the maintainers in CODEOWNERS or open a discussion. We'd rather you ask early than burn time the wrong way.
