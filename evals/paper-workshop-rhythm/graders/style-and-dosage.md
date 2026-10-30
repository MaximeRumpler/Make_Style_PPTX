---
type: llm
weight: 3
---

The agent ran `scripts/restyle.py` with `--style paper` and at least one arrangement (`--arrangement` with stagger, cascade, arc or deconstructed), limiting the arranged slides to about two or three (with `--slides` chosen for their content, or `--intensity balanced` or `calm`). It explains in plain words which slides were arranged and why they suit it. Fail if every slide was arranged (`--intensity all` or `bold` without reason), if no arrangement was used, or if deck.pptx was overwritten.
