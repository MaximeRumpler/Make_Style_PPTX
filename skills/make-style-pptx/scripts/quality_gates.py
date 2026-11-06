#!/usr/bin/env python3
"""Quality gates for a styled deck: warnings the agent turns into to-dos, never edits.

    quality_gates.py deck.pptx [--style paper] [--index layouts.csv,visuals.csv] [--strict] [--json FILE]

Five static gates, run on every deck the restyler writes (report key ``quality``) and on its own:

1. overlap: two text boxes that cover each other by more than a tenth of the smaller one;
2. contrast: each text run against the stack behind it (background, then every translucent
   or gradient shape drawn under it, ambience motif and giant word included, judged by the
   whole box, minus a soft edge, so it is conservative); below 4.5:1 (3:1 for large text), and worse than the
   same text on the bare background, is a warning. Text that was already faint is the content's
   business, not the style's;
3. reading order: the title comes before the body text (a kicker or number may lead it); ambience, giant words and
   decorations are flagged decorative;
4. effects budget: distinct kinds of effect on one slide (shadow, rotation, transparency,
   gradient, soft edge or glow, decorative group), against a budget (4 by default); gradients
   and glows that the style does not declare are flagged;
5. dosage: expressive slides (orbital, signature, giant word) in a row, or more than about
   3 in 10.

Output is JSON ({"slides", "warnings": [{"slide", "gate", "shape", "message"}], "by_gate"});
the exit code is 0, or 1 with ``--strict`` when there is a warning, 2 when the file cannot be read.
"""
from __future__ import annotations

import sys as _sys
from pathlib import Path as _Path

_sys.path.insert(0, str(_Path(__file__).resolve().parent))
import _env  # noqa: E402,F401  (switches to the private environment if needed)

import argparse
import json
import sys
from pathlib import Path
from typing import Dict, List, Optional

from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _ambience as AMB  # noqa: E402
import _arrange as AR  # noqa: E402

A, P = AMB.A, AMB.P
DECORATIVE_NAMES = ("ambience#", "giant#", "decoration#")
LEAD_IN = ("kicker", "eyebrow", "number", "label", "logo")  # small lines that may be read before the title
OVERLAP = 0.10  # share of the smaller text box
BUDGET = 4
# effects a style declares through its ambience motif; anything else is undeclared
DECLARED = {"mist": {"transparency", "soft edge"}, "wash": {"transparency", "soft edge"},
            "rain": {"transparency"}, "grid": {"transparency"}, "leak": {"transparency", "gradient"}}


def _is_decor(sh) -> bool:
    return sh.name.startswith(DECORATIVE_NAMES)


def _flagged_decorative(sh) -> bool:
    return sh._element.find(f".//{{{AMB.ADEC}}}decorative") is not None


def _text_shapes(slide):
    return [sh for sh in AMB._flat(slide.shapes) if AMB._text(sh).strip() and not _is_decor(sh)]


# ------------------------------------------------------------------ 1. overlap
def overlap_gate(slide, add) -> None:
    shapes = _text_shapes(slide)
    for i, a in enumerate(shapes):
        for b in shapes[i + 1:]:
            ba, bb = AMB._box(a), AMB._box(b)
            small = min(ba[2] * ba[3], bb[2] * bb[3])
            if small and AMB._overlap(ba, bb) / small > OVERLAP:
                add("overlap", a.name, f"text boxes '{a.name}' and '{b.name}' overlap")


# ------------------------------------------------------------------ 2. contrast
def _under(slide, theme) -> List[dict]:
    """Translucent or gradient shapes of the slide with their box, colour, alpha and z-index."""
    out = []
    for z, sh in enumerate(AMB._flat(slide.shapes)):
        if AMB._text(sh).strip():
            continue
        sppr = sh._element.find(P + "spPr")
        if sppr is None:
            continue
        fill = sppr.find(A + "solidFill")
        grad = sppr.find(A + "gradFill")
        el = fill[0] if fill is not None and len(fill) else None
        if el is None and grad is not None:
            stops = grad.findall(f"{A}gsLst/{A}gs")
            el = stops[0][0] if stops and len(stops[0]) else None  # strongest stop: conservative
        got = theme.resolve(el) if el is not None else None
        if got and got[1] < 0.99 and got[1] > 0:
            x, y, w, h = AMB._box(sh)
            soft = sppr.find(f"{A}effectLst/{A}softEdge")
            if soft is not None:  # the soft edge fades out: judge the solid core, like the ambience gate
                r = min(int(soft.get("rad", "0")), w // 2, h // 2)
                x, y, w, h = x + r, y + r, w - 2 * r, h - 2 * r
            out.append({"z": z, "box": (x, y, w, h), "rgb": got[0], "alpha": got[1]})
    return out


def contrast_gate(slide, add) -> None:
    theme = AMB.Theme(slide)
    bg = theme.background()
    layers = _under(slide, theme)
    if not layers:
        return
    flat = AMB._flat(slide.shapes)
    for run in AMB.text_runs(slide, theme):
        under = [lay for lay in layers if AMB._overlap(lay["box"], run["box"])]
        if not under:
            continue
        stack = bg
        for lay in under:  # in z order, each one composited over what lies below
            stack = AMB.blend(lay["rgb"], lay["alpha"], stack)
        base, got = AMB.contrast(run["rgb"], bg), AMB.contrast(run["rgb"], stack)
        need = run["need"]
        if got < need and got < base * (1 - AMB.WORSE):
            name = next((sh.name for sh in flat if AMB._box(sh) == run["box"] and AMB._text(sh).strip()), "text")
            add("contrast", name, f"text reads {got:.1f}:1 over the stack behind it (needs {need}:1, "
                                  f"{base:.1f}:1 on the bare background)")


# ------------------------------------------------------------------ 3. reading order
def order_gate(slide, add) -> None:
    for sh in slide.shapes:
        if sh.name.startswith(("ambience#", "giant#")) and not _flagged_decorative(sh):
            add("reading-order", sh.name, f"'{sh.name}' is decoration but not marked decorative")
    texts = [sh for sh in slide.shapes if AMB._text(sh).strip() and not _is_decor(sh)]
    titles = [sh for sh in texts if sh.name.lower().startswith("title")]
    if titles:
        before = [sh for sh in texts[:texts.index(titles[0])] if not sh.name.lower().startswith(LEAD_IN)]
        if before:
            add("reading-order", titles[0].name, f"'{before[0].name}' comes before the title in the reading order")


# ------------------------------------------------------------------ 4. effects budget
def effects(slide) -> Dict[str, int]:
    kinds: Dict[str, int] = {}

    def hit(kind):
        kinds[kind] = kinds.get(kind, 0) + 1

    for el in slide._element.iter():
        tag = el.tag
        if tag == A + "outerShdw" or tag == A + "innerShdw":
            hit("shadow")
        elif tag == A + "glow":
            hit("glow")
        elif tag == A + "softEdge":
            hit("soft edge")
        elif tag == A + "gradFill":
            hit("gradient")
        elif tag == A + "alpha" and int(el.get("val", "100000")) < 100000:
            hit("transparency")
        elif tag == A + "xfrm" and int(el.get("rot", "0")) % 21600000:
            hit("rotation")
    if any(_is_decor(sh) for sh in slide.shapes):
        hit("decorative group")
    return kinds


def effects_gate(slide, add, declared: set, budget: int) -> None:
    kinds = effects(slide)
    if len(kinds) > budget:
        add("effects", None, f"{len(kinds)} kinds of effect on one slide ({', '.join(sorted(kinds))}); "
                             f"the budget is {budget}")
    for kind in ("gradient", "glow", "soft edge"):
        if kind in kinds and kind not in declared:
            add("effects", None, f"{kind} is not declared by the style")


# ------------------------------------------------------------------ 5. dosage
def expressive_slides(prs, rows: Dict[str, dict]) -> List[int]:
    out = []
    for n, slide in enumerate(prs.slides, 1):
        lid = AR.slide_layout_id(slide)
        if AR.is_expressive(lid, rows.get(lid)) or any(sh.name.startswith("giant#") for sh in slide.shapes):
            out.append(n)
    return out


def dosage_gate(prs, rows, warn) -> None:
    hits = expressive_slides(prs, rows)
    for a, b in zip(hits, hits[1:]):
        if b == a + 1:
            warn(b, "dosage", None, f"expressive slides {a} and {b} follow each other")
    if len(hits) > max(1, round(len(prs.slides) * 0.3)):
        warn(0, "dosage", None, f"{len(hits)} expressive slides out of {len(prs.slides)} "
                                f"(slides {', '.join(map(str, hits))}); about 3 in 10 is the usual ceiling")


# ------------------------------------------------------------------ entry points
def check(prs, style: Optional[dict] = None, rows: Optional[Dict[str, dict]] = None, budget: int = BUDGET) -> dict:
    """The warnings of every gate on a deck. ``style`` (a loaded style) tells which effects are declared."""
    warnings: List[dict] = []
    amb = (style or {}).get("ambience") or {}
    declared = set(DECLARED.get(amb.get("motif"), set()))
    if (style or {}).get("giant"):
        declared.add("transparency")
    for n, slide in enumerate(prs.slides, 1):
        def add(gate, shape, message, n=n):
            warnings.append({"slide": n, "gate": gate, "shape": shape, "message": message})
        overlap_gate(slide, add)
        contrast_gate(slide, add)
        order_gate(slide, add)
        effects_gate(slide, add, declared, budget)
    dosage_gate(prs, rows or {}, lambda n, gate, shape, msg: warnings.append(
        {"slide": n, "gate": gate, "shape": shape, "message": msg}))
    by_gate: Dict[str, int] = {}
    for w in warnings:
        by_gate[w["gate"]] = by_gate.get(w["gate"], 0) + 1
    return {"slides": len(prs.slides), "warnings": warnings, "by_gate": dict(sorted(by_gate.items()))}


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("deck")
    ap.add_argument("--style", help="style name or file: its ambience and giant keys declare effects")
    ap.add_argument("--index", help="Make_Clean index file(s), comma separated, to know expressive layouts by id")
    ap.add_argument("--budget", type=int, default=BUDGET, help="kinds of effect allowed on one slide")
    ap.add_argument("--strict", action="store_true", help="exit 1 when there is a warning")
    ap.add_argument("--json", help="also write the report here")
    args = ap.parse_args(argv)
    try:
        import restyle as R
        style = R.load_style(args.style) if args.style else None
        rows: Dict[str, dict] = {}
        for path in (args.index or "").split(","):
            if path.strip():
                rows.update({r["id"]: r for r in R.read_index(Path(path.strip()))[1] if r.get("id")})
        report = check(Presentation(args.deck), style, rows, args.budget)
    except Exception as exc:  # noqa: BLE001
        print(json.dumps({"error": f"cannot read {args.deck}: {exc}"}))
        return 2
    if args.json:
        Path(args.json).write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, indent=2))
    return 1 if args.strict and report["warnings"] else 0


if __name__ == "__main__":
    sys.exit(main())
