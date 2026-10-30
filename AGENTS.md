# AGENTS.md

This repository ships one agent skill: **`skills/make-style-pptx/`**. It gives a clean deck another look without touching its content: other shapes for its role-tagged boxes, a few slides arranged off the plain grid, a quiet background motif, every change checked so the text still fits. It works on decks built by Make_Clean_PPTX and makes style packs for Make_Clean.

## When to use it

Whenever the user wants to restyle a deck, change its look or style, make it less boxy, softer, more organic or more hand-made, or see the same deck in another style (also: "change le style", "rends ce deck moins carré", "relooke ces slides").

## How to use it (harnesses without native skill loading)

1. Read `skills/make-style-pptx/SKILL.md` completely and follow its workflow. All paths in that file are relative to its own folder.
2. Read `skills/make-style-pptx/references/restyler.md` only to go beyond the two reference styles (custom style file, ambience, limits).
3. Use only the scripts in `skills/make-style-pptx/scripts/`. Do not move or redraw shapes by hand with python-pptx: the scripts check fit, overlap and reading order.
4. Requirements: Python 3.9+, `pip install -r skills/make-style-pptx/requirements.txt`. LibreOffice and poppler are optional (before/after images).

## Working on this repository

- Tests: `pip install -r requirements-dev.txt && pytest`. The round trip tests need a Make_Clean_PPTX checkout in `MAKE_CLEAN_DIR` (CI pins a release); without it they skip.
- The skill folder must stay **self-contained**: nothing in `skills/make-style-pptx/` may import from or read files outside it (people copy only that folder). Make_Clean_PPTX and Make_Colors_PPTX are companions, never dependencies.
- Scripts use the standard library plus python-pptx, Pillow and lxml only. Keep them Python 3.9 compatible.
- Every script prints machine-readable JSON (or clearly structured text) and uses distinct exit codes; keep that contract, the agent depends on it.
- Same input and same style give the same bytes: seeds come from the style and the slide, zip timestamps are fixed.
- Styles only use theme colour slots and theme fonts, so a brand theme still applies afterwards.
- `SKILL.md` stays under ~150 lines. Details go to `references/`, explained with the *why*, not as capitalised commands.
- Keep the README's design language: flat-square shields in `#10B981` on `#1C2333`, centred header, `* * *` separators, short bold-title tables.
