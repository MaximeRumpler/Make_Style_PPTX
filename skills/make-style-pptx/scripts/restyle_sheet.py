#!/usr/bin/env python3
"""Before/after contact sheet of a style on a few low-text layouts and visuals.

    restyle_sheet.py <library folder> --style organic [--ids L-24,L-38,...] [--png docs/restyle-organic.png]
    restyle_sheet.py <library folder> --style organic --arrangement stagger   # docs/arrange-stagger.png
    restyle_sheet.py deck.pptx --style paper --arrangement stagger,arc,deconstructed,cascade
        # docs/arrange-deck.png: deck built by Make_Clean from tests/fixtures/sample-deck/deck.json

Keeps only the chosen slides of the library (a small copy), restyles that copy, renders
both with LibreOffice and pdftoppm, and lays them out side by side, one row per
layout, with what the restyler did under each pair. With an arrangement, the right side
shows the arranged variant row (``L-38-stagger``) when there is one. With a deck, every
slide is shown, with the dosage decision under it. Needs LibreOffice for the PNG.
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parent))
import restyle as R  # noqa: E402
from contact_sheet import soffice  # noqa: E402

REPO = Path(__file__).resolve().parents[3]
DEFAULT_IDS = ["L-05", "L-24", "L-26", "L-28", "L-38", "L-47", "L-50", "V-PRO-04", "V-TIM-05"]
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def keep_slides(prs, numbers) -> None:
    """Drop every slide whose 1-based number is not in ``numbers`` (parts go with them)."""
    sld = prs.slides._sldIdLst
    for n, el in reversed(list(enumerate(list(sld), 1))):
        if n not in numbers:
            prs.part.drop_rel(el.get("{%s}id" % R_NS))
            sld.remove(el)


def subset(src: Path, dst: Path, ids) -> list:
    """Copy of a library folder with only the rows ``ids`` (slides renumbered).
    Returns the ids kept, in library order."""
    dst.mkdir(parents=True, exist_ok=True)
    kept = []
    for lib in ("layouts", "visuals"):
        if not (src / f"{lib}.csv").is_file():
            continue
        header, rows = R.read_index(src / f"{lib}.csv")
        rows = [r for r in rows if r["id"] in ids]
        if not rows:
            continue
        prs = Presentation(str(src / f"{lib}.pptx"))
        keep_slides(prs, {int(r["slide"]) for r in rows})
        rows.sort(key=lambda r: int(r["slide"]))
        for n, r in enumerate(rows, 1):
            r["slide"] = str(n)
            kept.append(r["id"])
        R.save_deterministic(prs, dst / f"{lib}.pptx")
        R.write_index(dst / f"{lib}.csv", header, rows)
    for extra in ("theme-contract.json", "style.json"):
        if (src / extra).is_file():
            shutil.copyfile(src / extra, dst / extra)
    return kept


def render_pages(pptx: Path, tmp: Path, dpi: int):
    from PIL import Image

    office = soffice()
    out = tmp / pptx.parent.name
    out.mkdir(parents=True, exist_ok=True)
    subprocess.run([office, "--headless", f"-env:UserInstallation=file://{tmp}/profile", "--convert-to", "pdf",
                    "--outdir", str(out), str(pptx)], check=True, capture_output=True, timeout=300)
    pdf = out / (pptx.stem + ".pdf")
    subprocess.run(["pdftoppm", "-png", "-r", str(dpi), str(pdf), str(out / pptx.stem)], check=True,
                   capture_output=True, timeout=300)
    return [Image.open(p).convert("RGB") for p in sorted(out.glob(pptx.stem + "-*.png"))]


def _font(size):
    from PIL import ImageFont

    for name in ("DejaVuSans.ttf", "LiberationSans-Regular.ttf", "Arial.ttf"):
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()


def caption(report: dict, row_id: str, shown: str = None) -> str:
    """What happened to ``row_id``; ``shown`` is the row on the right (its arranged variant)."""
    bits = []
    shown = shown or row_id
    for lib in report["libraries"].values():
        for it in lib["items"]:
            if it["role"] == "arrangement":
                if it["id"] != row_id or (shown != row_id and it.get("variant") != shown):
                    continue
                if shown == row_id and it["status"] == "arranged":
                    continue
            elif it["id"] != shown:
                continue
            if it["status"] in ("kept", "out-of-scope"):
                continue
            if it["status"] == "arranged":
                what = it["arrangement"]
            else:
                what = it.get("kind") if it["status"] == "restyled" else it["status"]
            bits.append(f"{it['shape'].split(' > ')[0]}: {what}")
    seen = []
    for b in bits:
        if b not in seen:
            seen.append(b)
    for lib in report["libraries"].values():
        for it in (lib.get("ambience") or {}).get("items", []):
            if it["id"] == row_id and it["status"] != "out-of-scope":
                n = it.get("dropped_by_contrast", 0)
                cut = f", {n} part{'s' if n > 1 else ''} cut by the contrast check" if n else ""
                seen.append(f"ambience: {it['motif']} {it['status']}{cut}")
    return ", ".join(seen[:6]) + (" ..." if len(seen) > 6 else "") or "nothing to restyle"


def _style(style_name: str, arrangement: str = None, intensity: str = None) -> dict:
    style = R.load_style(style_name)
    if arrangement:
        style["arrangement"] = [a.strip() for a in arrangement.split(",") if a.strip()]
    if intensity:
        style["intensity"] = intensity
    return style


def build(src: Path, style_name: str, ids, png: Path, dpi: int = 60, arrangement: str = None) -> Path:
    from PIL import Image, ImageDraw

    if not soffice() or not shutil.which("pdftoppm"):
        raise R.RestyleError("LibreOffice and pdftoppm are needed for the PNG")
    style = _style(style_name, arrangement)
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        before = tmp / "before"
        kept = subset(src, before, set(ids))
        after, report = R.run(before, style, tmp / "after")
        # the right side shows the arranged variant of a row when there is one
        slide_of = {}
        offset = 0
        for lib in ("layouts", "visuals"):
            if (after / f"{lib}.csv").is_file():
                _, rows = R.read_index(after / f"{lib}.csv")
                for r in rows:
                    slide_of[r["id"]] = offset + int(r["slide"]) - 1
                offset += len(rows)
        names = R.AR.style_arrangements(style)
        shown = {}
        for row_id in kept:
            shown[row_id] = next((f"{row_id}-{n}" for n in names if f"{row_id}-{n}" in slide_of), row_id)
        before_at = {row_id: i for i, row_id in enumerate(kept)}
        pages = {}
        for side, folder in (("before", before), ("after", after)):
            imgs = []
            for lib in ("layouts", "visuals"):
                if (folder / f"{lib}.pptx").is_file():
                    imgs += render_pages(folder / f"{lib}.pptx", tmp / side, dpi)
            pages[side] = imgs
    w, h = pages["before"][0].size
    pad, head, cap = 16, 56, 34
    sheet = Image.new("RGB", (pad * 3 + w * 2, head + len(kept) * (h + cap + pad) + pad), (236, 236, 232))
    d = ImageDraw.Draw(sheet)
    title = f"Restyler: {style['name']}" + (f" + {arrangement}" if arrangement else "")
    d.text((pad, 14), title, fill=(20, 20, 20), font=_font(22))
    d.text((pad + w + pad, 20), "before (left), after (right)",
           fill=(90, 90, 90), font=_font(13))
    y = head
    for row_id in kept:
        sheet.paste(pages["before"][before_at[row_id]], (pad, y))
        sheet.paste(pages["after"][slide_of[shown[row_id]]], (pad * 2 + w, y))
        label = row_id if shown[row_id] == row_id else f"{row_id} -> {shown[row_id]}"
        d.text((pad, y + h + 6), f"{label}  {caption(report, row_id, shown[row_id])}", fill=(40, 40, 40),
               font=_font(14))
        y += h + cap + pad
    png.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(png, optimize=True)
    return png


def build_deck(deck: Path, style_name: str, png: Path, dpi: int = 40, arrangement: str = None,
               intensity: str = None) -> Path:
    """Before/after of every slide of a deck, with the arrangement decision under each pair."""
    from PIL import Image, ImageDraw

    if not soffice() or not shutil.which("pdftoppm"):
        raise R.RestyleError("LibreOffice and pdftoppm are needed for the PNG")
    style = _style(style_name, arrangement, intensity)
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        (tmp / "before").mkdir()
        before = tmp / "before" / "deck.pptx"
        shutil.copyfile(deck, before)
        (tmp / "after").mkdir()
        after, report = R.run(before, style, tmp / "after" / "deck.pptx")
        pages = {side: render_pages(f, tmp / f"r-{side}", dpi) for side, f in (("before", before), ("after", after))}
    notes = {}
    for it in report["items"]:
        if it["role"] == "arrangement":
            notes[it["id"]] = (f"{it['arrange']}: {it['arrangement']}" if it["status"] == "arranged"
                               else f"{it['arrange']}: {it['status']}, {it.get('reason', '')}")
    w, h = pages["before"][0].size
    pad, head, cap = 12, 56, 26
    sheet = Image.new("RGB", (pad * 3 + w * 2, head + len(pages["before"]) * (h + cap + pad) + pad), (236, 236, 232))
    d = ImageDraw.Draw(sheet)
    hook = report["hooks"].get("arrangement") or {}
    d.text((pad, 12), f"Deck: {style['name']} + {', '.join(hook.get('requested', []))}", fill=(20, 20, 20),
           font=_font(20))
    d.text((pad, 36), f"intensity {hook.get('intensity')}: {hook.get('budget')} arranged slides at most, "
           f"{hook.get('gap')} slides apart at least. Before (left), after (right).", fill=(90, 90, 90), font=_font(13))
    y = head
    for i, (b, a) in enumerate(zip(pages["before"], pages["after"]), 1):
        sheet.paste(b, (pad, y))
        sheet.paste(a, (pad * 2 + w, y))
        text = notes.get(f"slide {i}", "no repeated units")
        d.text((pad, y + h + 5), f"slide {i}  {text}"[:150], fill=(40, 40, 40), font=_font(13))
        y += h + cap + pad
    png.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(png, optimize=True)
    return png


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("library", type=Path, help="library folder, or a built deck (.pptx)")
    ap.add_argument("--style", required=True)
    ap.add_argument("--ids", default=",".join(DEFAULT_IDS))
    ap.add_argument("--png", type=Path)
    ap.add_argument("--dpi", type=int, default=60)
    ap.add_argument("--arrangement", help="override the style's arrangements, comma separated "
                    "(stagger, cascade, arc, deconstructed)")
    ap.add_argument("--intensity", help="override the style's intensity (calm, balanced, bold), decks")
    args = ap.parse_args(argv)
    stem = f"arrange-{args.arrangement.split(',')[0]}" if args.arrangement else f"restyle-{Path(args.style).stem}"
    if args.library.suffix.lower() == ".pptx":
        stem = "arrange-deck"
    png = args.png or REPO / "docs" / f"{stem}.png"
    try:
        if args.library.suffix.lower() == ".pptx":
            build_deck(args.library, args.style, png, min(args.dpi, 40), args.arrangement, args.intensity)
            print(f"wrote {png}")
            return 0
        build(args.library, args.style, [i.strip() for i in args.ids.split(",") if i.strip()], png, args.dpi,
              args.arrangement)
    except R.RestyleError as exc:
        print(json.dumps({"error": str(exc)}))
        return 2
    print(f"wrote {png}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
