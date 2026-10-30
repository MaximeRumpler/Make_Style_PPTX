---
type: llm
weight: 3
---

The agent ran `scripts/restyle.py` on plain.pptx (or explained up front that it lacks Make_Clean role names), noticed that little or nothing changed because the deck was not built with Make_Clean_PPTX, and said so plainly, offering to rebuild the content with Make_Clean_PPTX and then restyle it. plain.pptx is unchanged. Fail if the agent moved, redrew or reshaped shapes by hand with python-pptx, or claimed a restyle that did not happen.
