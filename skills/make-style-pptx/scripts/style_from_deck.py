#!/usr/bin/env python3
"""Read the look of an example deck and draft a style file from it. Patterns only.

    style_from_deck.py example.pptx [--name mine] [--out styles/mine.json]
    style_from_deck.py --check styles/mine.json

The analysis keeps numbers and shape kinds, never text, names, pictures, colours or file names:
how many area shapes are rectangles, rounded, ellipses or free-form; how round the corners are;
how much is rotated, translucent, gradient, shadowed or outlined; how the left edges are
spread; the theme fonts. From these it writes a *draft* style (shapes per role, an arrangement,
an ambience motif) and one line of rationale per choice. The draft is a starting point for the
agent and the user to adjust: render it with ``restyle_sheet.py`` and change what is off.

``--check`` loads a style file through the restyler's own validation and prints ``{"ok": true}``
or the error (exit 2).
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Dict, List

from lxml import etree
from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _ambience as AMB  # noqa: E402

A, P = AMB.A, AMB.P
MIN_AREA = 0.3 * 914400 * 914400 * 0.3  # about 0.3 in x 0.3 in: smaller shapes are rules and marks
FILLED_OR_LINED = (A + "solidFill", A + "gradFill", A + "ln")


def _kind(sp) -> str:
    sppr = sp.find(P + "spPr")
    if sppr is None:
        return ""
    if sppr.find(A + "custGeom") is not None:
        return "freeform"
    geom = sppr.find(A + "prstGeom")
    prst = geom.get("prst") if geom is not None else ""
    if prst in ("rect", "snip1Rect", "snip2SameRect", "round1Rect"):
        return "rectangle"
    if prst in ("roundRect", "round2SameRect", "round2DiagRect"):
        return "rounded"
    if prst in ("ellipse", "donut", "pie", "chord"):
        return "ellipse"
    return "other" if prst and prst != "line" else ""


def _radius(sp) -> float:
    """Corner radius of a rounded rectangle, as a share of its shorter side (0 to 0.5)."""
    gd = sp.find(f"{P}spPr/{A}prstGeom/{A}avLst/{A}gd")
    if gd is not None and (gd.get("fmla") or "").startswith("val "):
        return min(0.5, int(gd.get("fmla")[4:]) / 100000)
    return 1 / 6  # the preset default, 16667


def analyse(prs) -> dict:
    kinds = {"rectangle": 0, "rounded": 0, "ellipse": 0, "freeform": 0, "other": 0}
    radii: List[float] = []
    total = rotated = translucent = gradient = shadow = outlined = filled = 0
    spread: List[float] = []
    chars = 0
    for slide in prs.slides:
        lefts = set()
        count = 0
        for sh in AMB._flat(slide.shapes):
            chars += len(AMB._text(sh))
            el = sh._element
            if el.tag != P + "sp" or (sh.width or 0) * (sh.height or 0) < MIN_AREA:
                continue
            sppr = el.find(P + "spPr")
            if sppr is None or not any(sppr.find(t) is not None for t in FILLED_OR_LINED):
                continue
            k = _kind(el)
            if not k:
                continue
            total += 1
            count += 1
            kinds[k] += 1
            lefts.add(round((sh.left or 0) / 91440))  # tenth of an inch
            if k == "rounded":
                radii.append(_radius(el))
            xfrm = sppr.find(A + "xfrm")
            rotated += xfrm is not None and int(xfrm.get("rot", "0")) % 21600000 != 0
            translucent += any(int(a.get("val", "100000")) < 100000 for a in sppr.iter(A + "alpha"))
            gradient += sppr.find(A + "gradFill") is not None
            shadow += sppr.find(f"{A}effectLst/{A}outerShdw") is not None
            ln = sppr.find(A + "ln")
            outlined += ln is not None and ln.find(A + "noFill") is None and len(ln) > 0
            filled += sppr.find(A + "solidFill") is not None or sppr.find(A + "gradFill") is not None
        if count > 1:
            spread.append(len(lefts) / count)
    share = lambda n: round(n / total, 3) if total else 0.0  # noqa: E731
    fonts = {}
    try:
        master = prs.slide_masters[0]
        root = etree.fromstring(master.part.part_related_by(AMB.RT_THEME).blob)
        for key, tag in (("heading", "majorFont"), ("body", "minorFont")):
            lat = root.find(f".//{A}fontScheme/{A}{tag}/{A}latin")
            if lat is not None:
                fonts[key] = lat.get("typeface")
    except (KeyError, etree.XMLSyntaxError, IndexError):
        pass
    radii.sort()
    return {"slides": len(prs.slides), "area_shapes": total,
            "kinds": {k: share(v) for k, v in kinds.items()},
            "corner_radius": round(radii[len(radii) // 2], 2) if radii else None,
            "rotated": share(rotated), "translucent": share(translucent), "gradient": share(gradient),
            "shadow": share(shadow), "outlined": share(outlined), "filled": share(filled),
            "left_edge_spread": round(sum(spread) / len(spread), 2) if spread else None,
            "chars_per_slide": round(chars / max(1, len(prs.slides))), "theme_fonts": fonts}


def draft(a: dict, name: str = "mine") -> dict:
    """A style file and the reasons for it, from an analysis."""
    why: List[str] = []
    k = a["kinds"]
    roles: Dict[str, dict] = {}
    if a["area_shapes"] < 3:
        why.append("too few area shapes to read a look from: describe the style in words instead")
    elif k["freeform"] >= 0.08 and a["outlined"] >= 0.1 and a["outlined"] >= a["filled"]:
        why.append(f"{k['freeform']:.0%} free-form and {a['outlined']:.0%} outlined, at least as much as filled: sketched, hand-made")
        roles = {"card": {"shapes": [{"kind": "handdrawn", "params": {"seed": "auto", "jitter": 0.02},
                                      "stroke": "tx2", "stroke_pt": 1}]},
                 "label": {"shapes": [{"kind": "handdrawn", "params": {"seed": "auto", "jitter": 0.03,
                                                                       "strokes": 1, "form": "auto"}}]}}
    elif k["freeform"] >= 0.08:
        why.append(f"{k['freeform']:.0%} free-form shapes: soft, organic silhouettes")
        roles = {"card": {"shapes": [{"kind": "blob", "params": {"seed": "auto", "wobble": 0.07}, "margin": 0.04},
                                     {"kind": "pebble", "params": {"seed": "auto", "wobble": 0.06}, "margin": 0.04},
                                     {"kind": "rounded", "params": {"radius": 0.3}}]},
                 "callout": {"shapes": [{"kind": "cloud", "params": {"lobes": 3}}, {"kind": "pill"}]},
                 "label": {"shapes": [{"kind": "circle"}, {"kind": "pill"}]}}
    elif k["rounded"] >= 0.3 and (a["corner_radius"] or 0) >= 0.35:
        why.append(f"{k['rounded']:.0%} rounded with a corner radius of {a['corner_radius']:.2f}: pill-like, friendly")
        roles = {"card": {"shapes": [{"kind": "rounded", "params": {"radius": 0.4}}]},
                 "callout": {"shapes": [{"kind": "pill"}]}, "label": {"shapes": [{"kind": "pill"}]}}
    elif k["rounded"] >= 0.3:
        why.append(f"{k['rounded']:.0%} rounded with a corner radius of {a['corner_radius']:.2f}: gently rounded")
        roles = {"card": {"shapes": [{"kind": "rounded", "params": {"radius": round(a['corner_radius'] or 0.15, 2)}}]},
                 "callout": {"shapes": [{"kind": "rounded", "params": {"radius": 0.25}}]}}
    elif k["ellipse"] >= 0.2:
        why.append(f"{k['ellipse']:.0%} ellipses and circles: round markers and figures")
        roles = {"keyfigure": {"shapes": [{"kind": "circle", "square": True, "align": "center"}]},
                 "label": {"shapes": [{"kind": "circle"}, {"kind": "pill"}]}}
    else:
        why.append(f"{k['rectangle']:.0%} plain rectangles: a sober, square look, the shapes stay as they are")
    arrangement = None
    if (a["rotated"] >= 0.1) or (a["left_edge_spread"] or 0) >= 0.7:
        arrangement = "stagger"
        why.append("rotated or loosely aligned shapes: items off the grid (stagger) on a few slides")
    ambience = None
    if a["gradient"] >= 0.15:
        ambience = {"motif": "leak", "density": "low"}
        why.append(f"{a['gradient']:.0%} gradients: a light leak behind covers")
    elif a["translucent"] >= 0.2:
        ambience = {"motif": "wash" if k["freeform"] >= 0.08 else "mist", "density": "low"}
        why.append(f"{a['translucent']:.0%} translucent shapes: a quiet wash behind covers")
    elif a["outlined"] >= 0.5 and a["filled"] < 0.4:
        ambience = {"motif": "grid", "density": "low"}
        why.append("mostly outlines: a fine grid behind covers, like squared paper")
    if a["shadow"] >= 0.3:
        why.append(f"{a['shadow']:.0%} of shapes carry a shadow: not reproduced, styles stay flat (shadows are an effect to declare)")
    style = {"name": re.sub(r"[^a-z0-9_-]+", "-", name.lower()).strip("-") or "mine",
             "description": "Draft from an example deck's patterns. Edit before use.",
             "seed": 1, "keep": 0.6, "max_chars": 320, "roles": roles, "arrangement": arrangement,
             "ambience": ambience}
    if a["theme_fonts"].get("heading") and a["theme_fonts"].get("body"):
        why.append(f"theme fonts {a['theme_fonts']['heading']} / {a['theme_fonts']['body']}: set `ambience.typography` "
                   "only if the fonts exist on the viewer's machine")
    return {"style": style, "rationale": why}


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("deck", nargs="?", help="example deck (.pptx)")
    ap.add_argument("--name", default="mine")
    ap.add_argument("--out", help="write the draft style here (a new file)")
    ap.add_argument("--check", metavar="STYLE", help="validate a style file and stop")
    args = ap.parse_args(argv)
    import restyle as R
    if args.check:
        try:
            style = R.load_style(args.check)
        except (R.RestyleError, OSError, ValueError) as exc:
            print(json.dumps({"ok": False, "error": str(exc)}))
            return 2
        print(json.dumps({"ok": True, "style": style["name"]}))
        return 0
    if not args.deck:
        ap.error("an example deck or --check is required")
    try:
        analysis = analyse(Presentation(args.deck))
    except Exception as exc:  # noqa: BLE001
        print(json.dumps({"error": f"cannot read {args.deck}: {exc}"}))
        return 2
    result = {"analysis": analysis, **draft(analysis, args.name)}
    if args.out:
        out = Path(args.out)
        if out.exists():
            print(json.dumps({"error": f"{out} already exists"}))
            return 2
        out.parent.mkdir(parents=True, exist_ok=True)
        out.write_text(json.dumps(result["style"], indent=2) + "\n", encoding="utf-8")
        try:
            R.load_style(str(out))
        except (R.RestyleError, ValueError) as exc:
            out.unlink()
            print(json.dumps({"error": f"the draft does not validate: {exc}"}))
            return 2
        result["written"] = str(out)
    print(json.dumps(result, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
