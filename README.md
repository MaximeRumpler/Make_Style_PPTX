<div align="center">

[![Make_Style_PPTX](docs/banner.svg)](https://github.com/MaximeRumpler/Make_Style_PPTX)

*From your content, first Make_Clean, then give it a style.*

[![License](https://img.shields.io/badge/License-EUPL--1.2-10B981?style=flat-square&labelColor=1C2333)](LICENSE)
[![Status](https://img.shields.io/badge/Status-in_preparation-10B981?style=flat-square&labelColor=1C2333)](#roadmap)
[![Companion](https://img.shields.io/badge/Companion_of-Make__Clean__PPTX-10B981?style=flat-square&labelColor=1C2333)](https://github.com/MaximeRumpler/Make_Clean_PPTX)

</div>

* * *

## What it will do

Make Style PPTX is a companion *skill* of [**Make_Clean_PPTX**](https://github.com/MaximeRumpler/Make_Clean_PPTX).

Make_Clean_PPTX gives your notes their **form**: one message per slide, a matching design and chart. Make Style PPTX gives the same deck a **look**: the shape of the cards, the corners, the typefaces, the decorative motifs, the background. Your content, your layout and your figures stay exactly as they are.

| **Editorial** | **Organic** | **Geometric** | **Tech** | **Soft** |
|:---:|:---:|:---:|:---:|:---:|
| serif titles, paper tone | rounded shapes, arcs | circles, clean cards | dark, outlined, crisp | soft shapes, light shadows |

You pick a style by name or by tone (a board meeting, a keynote, a workshop), or show it a deck you like and it builds a new style from it. A small set of signature graphics (gears, a light bulb, a puzzle, a winding road…) is there to give a deck a clear identity, used sparingly.

* * *

## The Make PPTX set

Part of [**Make PPTX**](https://github.com/MaximeRumpler/Make_PPTX), a set of three skills from your notes to a finished deck.

| **Make_Clean_PPTX** | **Make_Style_PPTX** | **Make_Colors_PPTX** |
|:---|:---|:---|
| **Form.** Notes in, clean deck out. | **Look.** The same deck in another style. | **Brand.** Your colours, fonts and logo. |
| [Install](https://github.com/MaximeRumpler/Make_Clean_PPTX#install) · [Gallery](https://maximerumpler.github.io/Make_Clean_PPTX/) | In preparation | In preparation |

Each one works on its own; together they take you from raw notes to a finished, branded deck. All of them work in Claude Code, Claude.ai, Codex and other assistants that read skills.

* * *

## Roadmap

In preparation. Follow [Make_Clean_PPTX](https://github.com/MaximeRumpler/Make_Clean_PPTX) for updates; the first release will be announced there.

First brick in place: the **shapes library** (`skills/make-style-pptx/scripts/_shapes.py`). Drops, clouds, blobs, pebbles, tickets, speech bubbles, torn paper, polygons, hand-drawn cards and filled icons become native, editable PowerPoint shapes drawn with smooth curves, each with the safe box where its text goes. See them all at three sizes in the [contact sheet](docs/shapes-contact-sheet.png) (rebuild it with `python skills/make-style-pptx/scripts/contact_sheet.py`).

Second brick: the **restyler** (`skills/make-style-pptx/scripts/restyle.py`). It turns a Make_Clean library into a style pack: cards, callouts, key figures and labels take a new shape, their text moves into the shape's safe box, and the character limits of the index are recomputed. When a text would no longer fit, the container keeps its standard shape. Two reference styles, *organic* and *paper*, before and after: [organic](docs/restyle-organic.png), [paper](docs/restyle-paper.png). Details in [references/restyler.md](skills/make-style-pptx/references/restyler.md).

Third brick: **arrangements** (`scripts/_arrange.py`). On the layouts whose units Make_Clean marks as rearrangeable (three columns, four figures, a vertical sequence…), a style can lay the units out as a stagger, a cascade, an arc or a deconstructed 12-column grid. Positions depend only on the arrangement and the number of units; when a text would overflow, overlap or leave the margins, the slide keeps its plain grid. Before and after: [stagger](docs/arrange-stagger.png), [arc](docs/arrange-arc.png), [cascade](docs/arrange-cascade.png), [deconstructed](docs/arrange-deconstructed.png).

Fourth brick: the **ambience** (`scripts/_ambience.py`). A style can put one quiet motif behind its covers, statements and dividers: a soft mist, thin rain streaks or a fine technical grid. The motifs are native, editable shapes in theme colours, sit behind the content, never go on dense slides, and any part that would make a text harder to read than WCAG's 4.5:1 is left out. *organic* uses the mist, *paper* the grid: [mist](docs/ambience-mist.png), [grid](docs/ambience-grid.png), [rain](docs/ambience-rain.png).

* * *

## License

[EUPL-1.2](LICENSE), the open source licence of the European Union: free to use; if you share an improved version, share it under the same licence. **The presentations and styles you make are yours**, with no obligation ([NOTICE](NOTICE)).

<div align="center">

Built by [**Maxime Rumpler**](https://github.com/MaximeRumpler) · [![KairosTech](https://img.shields.io/badge/Kairos-Tech-10B981?style=flat-square&labelColor=1C2333&logoColor=white)](https://github.com/KairosTech-fr) · [kairostech.fr](https://kairostech.fr)

<sub>With thanks to the creators and contributors of [impeccable](https://github.com/pbakaus/impeccable), whose design skills for AI agents inspired this set, and of Wikipedia's [Signs of AI writing](https://en.wikipedia.org/wiki/Wikipedia:Signs_of_AI_writing), which guides how its texts are written.</sub>

<p align="right">
  <img src="https://komarev.com/ghpvc/?username=MaximeRumpler&repo=Make_Style_PPTX&color=44cc11&style=flat-square&label=views" alt="Repository Views" />
</p>

</div>
