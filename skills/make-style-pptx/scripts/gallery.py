#!/usr/bin/env python3
"""The same deck in every style: a catalogue page and one contact sheet.

    gallery.py [deck.pptx] [--slides 1,3,5] [--styles organic,paper] [--out gallery] [--dpi 36]

Restyles one deck (the sample deck by default) in every style of ``styles/`` and writes, in
``--out`` (a new folder, default ``gallery`` in the current folder):

- ``gallery.md``: one table, a row per style (family, when to use it, light or dark theme,
  what it changes), from ``assets/catalog.json`` and the style files;
- ``gallery.png``: one row per style with the chosen slides side by side; the second one is
  replaced by the first slide the style arranged off the grid, when it arranges any (needs
  LibreOffice and pdftoppm; without them only the table is written);
- ``gallery.json``: per style the restyle summary and the quality warnings, so the run doubles
  as a regression check (a style must not break its own deck).

Every style is judged on the same deck, so the sheet shows what the style does, not what the
content happens to be.
"""
from __future__ import annotations

import argparse
import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import restyle as R  # noqa: E402
import restyle_sheet as RS  # noqa: E402

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent
CATALOG = SKILL / "assets" / "catalog.json"
SAMPLE = SKILL / "assets" / "sample-deck.pptx"
FAMILIES = ("Sober", "Organic", "Tech", "Expressive", "Mood")
DEFAULT_SLIDES = (1, 3, 5)


def catalog() -> dict:
    return json.loads(CATALOG.read_text(encoding="utf-8"))


def ordered(names, cat) -> list:
    return sorted(names, key=lambda n: (FAMILIES.index(cat[n]["family"]) if n in cat else 99, n))


def layers(style: dict) -> str:
    """What the style changes, in a few words."""
    parts = []
    shapes = sorted({c["kind"] for r in style["roles"].values() for c in r.get("shapes", [])})
    if shapes:
        parts.append("shapes: " + ", ".join(shapes))
    arr = style.get("arrangement")
    if arr:
        parts.append("arrangement: " + (arr if isinstance(arr, str) else ", ".join(arr if isinstance(arr, list) else [str(arr)])))
    if (style.get("ambience") or {}).get("motif"):
        parts.append("ambience: " + style["ambience"]["motif"])
    if style.get("giant"):
        parts.append("giant word: " + style["giant"]["mode"])
    return "; ".join(parts) or "nothing: the baseline"


def table(names, cat) -> str:
    rows = ["| Style | Family | Use when | Theme | What it changes |", "|---|---|---|---|---|"]
    for n in names:
        st = R.load_style(n)
        c = cat.get(n, {"family": "", "use_when": st.get("description", ""), "mode": ""})
        rows.append(f"| `{n}` | {c['family']} | {c['use_when']} | {c['mode']} | {layers(st)} |")
    return "\n".join(rows) + "\n"


def run_style(deck: Path, name: str, tmp: Path) -> tuple:
    out, report = R.run(deck, R.load_style(name), tmp / f"{name}.pptx", force=True)
    arranged = (report["hooks"].get("arrangement") or {}).get("arranged", [])
    return out, {"summary": report["summary"]["total"], "quality": report.get("quality", {}).get("warnings", []),
                 "arranged": [int(a.split()[1]) for a in arranged]}


def build(deck: Path, names, slides, out: Path, dpi: int) -> dict:
    cat = catalog()
    names = ordered(names, cat)
    out.mkdir(parents=True, exist_ok=True)
    (out / "gallery.md").write_text("# Every style on the same deck\n\n" + table(names, cat), encoding="utf-8")
    results, images = {}, {}
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        for name in names:
            pptx, results[name] = run_style(deck, name, tmp)
            try:
                pages = RS.render_pages(pptx, tmp, dpi)
                shown = list(slides)
                if results[name]["arranged"] and len(shown) > 1:  # show a slide the style rearranged
                    shown[1] = results[name]["arranged"][0]
                results[name]["shown"] = shown
                images[name] = [pages[n - 1] for n in shown if n <= len(pages)]
            except Exception as exc:  # noqa: BLE001 (no LibreOffice or poppler: the table is still written)
                results[name]["render"] = f"skipped: {exc.__class__.__name__}"
        if images:
            _sheet(images, names, out / "gallery.png")
    (out / "gallery.json").write_text(json.dumps(results, indent=2) + "\n", encoding="utf-8")
    return results


def _sheet(images: dict, names, png: Path) -> None:
    from PIL import Image, ImageDraw

    first = next(iter(images.values()))[0]
    w, h = first.size
    n = max(len(v) for v in images.values())
    label, gap = 150, 8
    sheet = Image.new("RGB", (label + n * (w + gap) + gap, len(names) * (h + gap) + gap), "white")
    draw, font = ImageDraw.Draw(sheet), RS._font(14)
    for r, name in enumerate(names):
        y = gap + r * (h + gap)
        draw.text((gap, y + h // 2 - 8), name, fill=(30, 30, 30), font=font)
        for c, im in enumerate(images.get(name, [])):
            sheet.paste(im, (label + gap + c * (w + gap), y))
            draw.rectangle([label + gap + c * (w + gap) - 1, y - 1, label + gap + c * (w + gap) + w, y + h],
                           outline=(200, 200, 200))
    png.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(png, optimize=True)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("deck", nargs="?", type=Path, default=SAMPLE)
    ap.add_argument("--slides", default=",".join(map(str, DEFAULT_SLIDES)), help="slides to show (numbers)")
    ap.add_argument("--styles", help="only these styles, comma separated (default: all)")
    ap.add_argument("--out", type=Path, default=Path("gallery"))
    ap.add_argument("--dpi", type=int, default=36)
    args = ap.parse_args(argv)
    known = sorted(p.stem for p in R.STYLES_DIR.glob("*.json"))
    names = [s.strip() for s in args.styles.split(",") if s.strip()] if args.styles else known
    try:
        slides = [int(n) for n in args.slides.split(",") if n.strip()]
        for n in names:
            R.load_style(n)
        results = build(args.deck, names, slides, args.out, args.dpi)
    except (R.RestyleError, ValueError, OSError) as exc:
        print(json.dumps({"error": str(exc)}))
        return 2
    warned = {n: len(r["quality"]) for n, r in results.items() if r["quality"]}
    print(json.dumps({"out": str(args.out), "styles": len(results), "with_quality_warnings": warned}, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
