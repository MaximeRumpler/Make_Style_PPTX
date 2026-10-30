---
type: llm
weight: 2
---

Before answering, the agent checked the result: it ran `scripts/restyle_sheet.py` with `--png` (or rendered the output another way) and looked at the image, or, if rendering was impossible, read the JSON report and said the visual check was skipped. Fail if the agent delivered without any check.
