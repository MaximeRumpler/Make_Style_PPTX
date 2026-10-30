# Evals

Behaviour checks for the skill, run with Claude Code's plugin eval runner:

```bash
claude plugin eval . --allow-tools "Bash(python3:*)" "Bash(cp:*)" Write Edit --runs 1 --no-publish
```

Each case is a realistic request (`prompt.md`) and graders that judge the result (`graders/*.md`). Most cases start from the sample deck bundled in the skill (`assets/sample-deck.pptx`, a Make_Clean deck). They cover what unit tests cannot: does the agent pick a style that fits the request, dose the off-grid slides to the audience, check the result, keep the user's text and leave the user's file alone. Every run restyles real decks, so it costs a few agent runs; run them before a release, not on every commit.
