@AGENTS.md

## Claude Code notes

- The repository is also a Claude Code plugin: `.claude-plugin/` (manifests) and `skills/` (the skill).
- After editing the skill, validate it with the skill-creator checker if available: `python -m scripts.quick_validate skills/make-style-pptx`.
- Do not commit generated decks, packs or renders (`build/`, `*.report.json` are ignored). Do commit the `docs/` images a change regenerates.
