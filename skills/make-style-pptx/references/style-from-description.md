# A style from a description or an example deck

When none of the reference styles fits ("rain and sadness", "like a children's book", "serious but warm", or "like this deck"), build a style file. A style is three layers; decide each one and say why in a line.

## 1. From a description: map the mood to the layers

| Mood | Shapes (`kind`) | Arrangement | Ambience |
|---|---|---|---|
| soft, warm, human | `blob`, `pebble`, `cloud`, `pill`, `circle` | `stagger` or none | `mist`, `wash` |
| hand-made, drafts, workshop | `handdrawn`, `torn`, `ticket` | `stagger`, `cascade` | `grid` |
| sad, calm, rainy | `drop` (short labels only), `cloud`, `pebble` | `cascade` (falling) | `rain`, `mist` |
| editorial, bold, magazine | `rectangle`, `polygon`, `rounded` with small radius | `deconstructed` | none, or a giant word (`--giant`) |
| warm light, nostalgic | `rounded`, `pebble` | `arc` | `leak`, `wash` |
| technical, precise | `rectangle`, `rounded` (0.1) | none | `grid` (`"marks": "cross"`) |

Say the rationale in plain words: "rain and sadness: drops and clouds for the callouts, items falling down the slide, thin rain behind the covers". If two moods pull apart, choose the quieter one for the shapes and let the ambience carry the rest.

## 2. From an example deck: patterns only

```bash
python3 scripts/style_from_deck.py example.pptx --name mine --out styles/mine.json
```

The script reads numbers and shape kinds only (share of rectangles, rounded, ellipses and free-form shapes; corner radius; rotation, transparency, gradient, shadow and outline shares; how spread the left edges are; theme fonts) and drafts a style with one line of rationale per choice. It keeps **no text, no shape or file names, no pictures and no colours**: the example is inspiration, nothing is copied. A deck with few area shapes cannot be read; describe the look in words instead. Shadows are not reproduced (styles stay flat).

## 3. Check, render, adjust

1. `python3 scripts/style_from_deck.py --check styles/mine.json` validates the file with the restyler's own rules (unknown shapes, roles or motifs are refused).
2. Render the before/after on a sample deck: `python3 scripts/restyle_sheet.py assets/sample-deck.pptx --style styles/mine.json --png build/mine.png`, and look at it.
3. Adjust one thing at a time (a shape order, a radius, the motif), render again. **Stop after two rounds** with the user and keep the best version.
4. Colours come from the theme (Make_Colors); never write a colour into a style.
