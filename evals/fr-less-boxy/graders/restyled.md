---
type: llm
weight: 3
---

The agent used the make-style-pptx skill and its `scripts/restyle.py` on deck.pptx with the `organic` style (the soft, warm one), wrote a new file (for example deck-organic.pptx) and left deck.pptx unchanged. Its final answer is in French, short (style used, what changed, anything kept plain and why) and gives the path of the new file. Fail if shapes were moved or redrawn by hand with python-pptx, if deck.pptx was overwritten, or if no restyled file exists.
