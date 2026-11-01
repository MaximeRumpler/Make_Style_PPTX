# Quality gates

`scripts/quality_gates.py` runs on every deck the restyler writes (the report's `quality` key) and on its own (`python3 scripts/quality_gates.py deck.pptx --style paper`). It only warns: each warning is a to-do for the agent, never an automatic edit. `--strict` exits 1 when there is a warning.

| Gate | What it flags | What to do |
|---|---|---|
| `overlap` | two text boxes covering each other by more than a tenth of the smaller one | pick another shape or arrangement for that slide, or shorten the text |
| `contrast` | a text run below 4.5:1 (3:1 for large text) over the stack behind it, and worse than on the bare background | drop or lighten the motif or the giant word on that slide; text that was already faint is the content's own matter |
| `reading-order` | body text before the title in the XML; ambience or giant word not marked decorative | move the title first; mark decoration decorative |
| `effects` | more than 4 kinds of effect on one slide (shadow, rotation, transparency, gradient, glow, soft edge, decorative group), or a gradient, glow or soft edge the style does not declare | remove the extra effect; declare it in the style if it is meant |
| `dosage` | expressive slides (orbital, signature, giant word) in a row, or more than about 3 in 10 | spread them out or drop some; keep more only if the request asks for it |

Why a warning list and not a hard stop: the amount of expression depends on the request and the audience (the dosage is a soft default), but a text that cannot be read or two boxes on top of each other are never what anyone wants.

The contrast gate is conservative: it uses the whole box of every translucent shape under a run (minus a soft edge), where the ambience layer's own check looks at the solid core, so a small disagreement means the text sits at the edge of a motif.

## Review checklist (by eye, on the rendered contact sheet)

1. **Meaning**: every decorative element must answer "what does it say?" or go. No random rotation, no gradient for its own sake, no texture without mood, no made-up data, no lorem ipsum.
2. **Giant word**: it comes from the slide's own message, and the slide still reads without it.
3. **Reference**: put the result next to the style's reference sources when a style is built from examples; keep what the examples do, not their content.
4. **Limits**: render the slides with the shortest and the longest realistic text before shipping a new arrangement or shape.

Not automated yet: rendering every composition at minimum and maximum content in CI, and sampling contrast from the render instead of the XML.
