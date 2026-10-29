# Restyler

`scripts/restyle.py` swaps the geometry of the role-tagged containers of a Make_Clean library or deck for shapes of the shapes library (`_shapes.py`). It works on the interface described in Make_Clean's `references/style-contract.md` (role tags, `role=` flags, per-pack limits) and needs only python-pptx, lxml and Pillow.

```
python scripts/restyle.py <library folder> --style organic [--out <pack folder>]
python scripts/restyle.py deck.pptx --style paper [--out deck-paper.pptx]
python scripts/restyle_sheet.py <library folder> --style organic     # before/after PNG in docs/
```

The input is never written. A library gives a **style pack** folder that Make_Clean builds with as is (`build_deck.py --local-dir <pack>`):

| File | What it holds |
|---|---|
| `layouts.pptx`, `visuals.pptx` | the restyled libraries: same slides, same ids, same shape names |
| `layouts.csv`, `visuals.csv` | the indexes, byte-identical except the lowered `limit` of the fields that moved into a smaller text box |
| `theme-contract.json`, `style.json` | copied from the input when present |
| `restyle.json` | the style definition used |
| `restyle-report.json` | every container: restyled (with its shape, parameters, new limits and word limits), fallback (with the reason of each candidate), text-heavy, out of scope or kept |

A deck gives `<deck>-<style>.pptx` and `<deck>-<style>.report.json`.

## Style definition

A JSON file in `styles/` (or anywhere, passed by path):

```json
{
  "name": "organic",
  "seed": 11,
  "keep": 0.6,
  "max_chars": 320,
  "roles": {
    "card": {"shapes": [
      {"kind": "blob", "params": {"seed": "auto", "wobble": 0.07}, "margin": 0.04},
      {"kind": "pebble", "params": {"seed": "auto", "wobble": 0.06}, "margin": 0.04},
      {"kind": "rounded", "params": {"radius": 0.3}}
    ]},
    "keyfigure": {"shapes": [
      {"kind": "circle", "square": true, "align": "center"},
      {"kind": "blob", "params": {"seed": "auto", "wobble": 0.05}, "square": true, "align": "center"}
    ]}
  },
  "arrangement": null,
  "ambience": {"motif": "mist", "density": "low"}
}
```

- `roles`: for `card`, `callout`, `keyfigure`, `label`, `panel` and `decoration`, the candidate shapes in order of preference. A role that is missing is left alone. `logo` and `frame` are refused: they are never restyled.
- Candidate keys: `kind` (a shapes library kind) and `params` (its parameters; `"seed": "auto"` gives each container its own stable seed, `"edges": "auto"` tears the edges of a panel that face the slide, `"form": "auto"` draws a hand-drawn diamond where the standard shape is a diamond, a hand-drawn card elsewhere); `keep` (overrides the style's; no reference style goes below 0.6); `margin` (room kept around the shape, share of the shorter side, so pebbles do not touch their neighbours); `align: "center"` (centres the text it holds); `square` (shrinks the box towards a square first, never below `min_cover` of a side, so a key figure becomes a disc, not a tall oval); `stroke` and `stroke_pt` (line colour as a theme colour name and width, for hand-drawn shapes whose pencil strokes need a line); `any_aspect` (draws the shape at the box's proportions, for torn bands of any length).
- `keep`: the least share of a field's character limit that the new text box must keep (0.6 = a 100-character field may drop to 60, not below).
- `max_chars`: containers whose fields add up to more characters are text-heavy and keep their standard shape. A role can set its own.
- `scope`: which rows are restyled (default: layouts of density `Low` or `Medium`, never `Table` or `Frame`; visuals of kind `diagram` or `timeline`, never charts, tables, matrices or illustrations).
- `seed`, `min_cover` (how much of the box a shape must keep when shrunk to its proportions, default 0.85), `padding` (share of the container's original inner margin kept around the text, default 0.75), `min_chars_per_line` (default 10).
- `ambience`: a background motif behind low-text slides, see [Ambience](#ambience). `null` (the default) adds nothing.
- `arrangement`: the later layer (#1607). Accepted and recorded in the report as not applied.

## What happens to one container

1. The container's box is shrunk (centred) to the shape's aspect bounds when needed; the shape is built there and its safe text box computed, pulled in to keep the container's original inner margin.
2. The text fields inside the container (and the rules or ornaments among them) are moved into the safe box: untouched when they already fit, else shifted the least possible; when too tall, a vertical stack first gives up its gaps and the empty room under each text.
3. Each field is checked at its own font size: at least one line, no word broken, the template text (in a deck: the real text) still fits, at least `min_chars_per_line` per line, and the new character limit (the style contract formula) keeps `keep` of the old one. Fonts are never changed.
4. In a round silhouette (`blob`, `pebble`, `drop`, `circle`, `cloud`, `ellipse`) the fields are then centred as a group, horizontally and vertically, in the safe box: a stack first gives up the height its character limits do not need (same limits, same word counts) and closes up, keeping its gaps; a single field is also anchored in the middle, and a container's own text is centred. Rectangles, rounded and hand-drawn cards, tickets and torn paper keep the original top-left layout.
5. Only when every field passes is the geometry swapped (`a:custGeom`, fill, line and theme colours unchanged) and the fields moved; otherwise the next candidate is tried. When none passes, the container keeps its standard shape and the report says why.

Same input and same style give the same bytes (fixed zip timestamps, seeds from the layout id and shape name).

## Ambience

`scripts/_ambience.py` runs after the shapes, on the same copy. It puts **one** motif behind the content of low-text slides:

| Motif | What it draws |
|---|---|
| `mist` | three soft, overlapping blobs (soft-edge effect) gathered off the emptiest corner |
| `rain` | thin streaks at one slant on the emptier side, thinning out towards the content, never touching text, rules or pictures |
| `grid` | a fine grid on the emptier side, aligned to the slide edge, fading out towards the content; `"marks": "cross"` draws small crosses at the intersections instead of lines |

```json
"ambience": {"motif": "grid", "density": "low", "marks": "lines"}
```

- Native shapes only (custom geometry, theme colours with a transparency): no bitmap, every part can be edited or deleted. The motif is one group, `ambience#1 > <motif>`, first in the shape tree (behind everything), marked decorative so screen readers skip it. `ambience#1` is not a role tag: the restyler and Make_Clean ignore it.
- Colours: `colour` on light slides (default `tx2`), `on_dark` on dark ones (default `bg1`), any theme slot. The slide background is read from the slide, its layout or its master, so a new theme (Make_Colors) rebrands the motif.
- `density`: `low` (default) or `medium`, the amount and opacity of the motif.
- Where: in a library, layouts of the `categories` (default `Cover`, `Statement`, `Divider`) whose density is in `layout_density` (default `Low`); never tables, frames or visuals. In a deck, slides with no chart, table or picture, at most `max_fields` text fields (8) and `max_chars` characters (280), and some text at `min_pt` (28 pt) or more.
- `seed`: added to the style seed; each slide also gets a stable seed from its id, so the same input gives the same bytes.
- **Contrast gate**: every text run over a motif part is checked against the colour the part leaves under it (WCAG 2.x: 4.5:1, 3:1 for 18 pt or 14 pt bold and above). A part that would take a run below its threshold is left out; a run that is already below it without any motif may lose at most 3 % of its contrast. Overlapping mist blobs add up; lines are judged one by one, so only the grid cells under a failing run disappear. Text on its own fill or over an opaque shape is not checked (the motif cannot be seen there). When nothing is left, the slide keeps no motif and the report says so.
- `typography` (optional): `{"major": "<font>", "minor": "<font>"}` sets the theme heading and body fonts of the pack or deck. Character limits are not recomputed: choose fonts close to the original widths.
- The report gets an `ambience` entry per library (or deck): for each slide its status (`applied`, `skipped`, `out-of-scope`), motif, colour slot, number of parts and parts cut by the contrast check.

Paper grain is not offered: PowerPoint's preset pattern fills render as hatching in LibreOffice, and dots drawn one by one would weigh hundreds of kilobytes per slide.

## Limits

- **Word limits** are computed for every moved field and written in `restyle-report.json` (`words`). The index schema of Make_Clean has no word-limit flag yet, so they are not in the CSV.
- In round shapes the template text is centred with the room its limit gives it: a field whose limit allows four lines but holds two in the template sits slightly above the middle until it is filled.
- Wide bands (more than about 3:1 once the 85 % rule is applied, e.g. the stage bars of V-PRO-04) are too long for a pebble and stay `rounded` in organic.
- Fields inside the container are found by geometry (their box inside the container's box). Containers nested in another (a card in a lane) are never moved; the outer shape falls back when they would stick out.
- The contrast gate works on text boxes, not on the glyphs: a motif part that crosses the empty end of a wide text box is judged as if it sat under the text.
- In a deck, roles come from the shape names only: fields that are also containers (`role=` flags in the index) are not recognised, and slides with a chart or a table are left alone.
