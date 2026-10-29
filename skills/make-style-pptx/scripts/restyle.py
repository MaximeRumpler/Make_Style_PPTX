#!/usr/bin/env python3
"""Restyle a Make_Clean library or a built deck: swap the geometry of role-tagged containers.

    restyle.py <library folder | deck.pptx> --style organic [--out DIR|FILE] [--report FILE] [--force]
               [--arrangement stagger,arc] [--intensity calm|balanced|bold|all] [--slides 2,6,9]

Input
    a library folder (layouts.pptx/.csv and visuals.pptx/.csv, a Make_Clean pack or
    the bundled ``assets/``), or a deck built by Make_Clean (``.pptx``).
Style
    a name from ``styles/`` (``organic``, ``paper``) or a ``.json`` file: for each role
    (card, callout, keyfigure, label, panel, decoration) a list of candidate shapes
    from the shapes library, tried in order.
Output (always a copy, the input is never written)
    library: a style pack folder (default ``<input>-<style>/``) with layouts/visuals
    pptx + csv (same ids, same field names, limits recomputed for the new text boxes),
    the input's theme-contract.json and style.json, ``restyle.json`` (the style used)
    and ``restyle-report.json``. deck: ``<deck>-<style>.pptx`` plus its report.

For each container in scope the restyler builds the new shape at the container's box
(shrunk to the shape's aspect bounds when needed), moves the text fields it holds into
the shape's safe text box, and checks every one of them at its own font size: at least
one line, no word broken, the template text (deck: the real text) still fits, and the
new character limit keeps at least ``keep`` of the old one. Fonts are never changed.
When no candidate passes, the container keeps its standard shape and the report says
why. Same input and style give the same bytes.

Scope (low-text slides by default, ``scope`` in the style): layouts whose density is
listed and whose category is not excluded (tables and frames never), visuals of the
listed kinds (diagrams and timelines; charts, tables, matrices and illustrations
never), and only containers whose fields hold at most ``max_chars`` characters in all.
The ``logo`` role and ``frame`` role are never touched.

Arrangement (#1607, ``_arrange.py``): when the style names one or more, a library gets an
arranged copy of each layout tagged in the index's ``arrange`` column, as a new opt-in row
(``L-38-stagger``, tagged ``expressive``); a deck gets some slides arranged in place, as
many as the ``intensity`` preset or the ``dosage`` (or ``--slides``) asks. Units move
before the shapes are swapped. Ambience
(``ambience`` in the style, see ``_ambience.py``): a quiet background motif behind
covers, statements and dividers, after the shapes.
"""
from __future__ import annotations

import argparse
import copy
import csv
import io
import json
import math
import re
import shutil
import sys
import zipfile
import zlib
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from lxml import etree
from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _ambience as AMB  # noqa: E402
import _arrange as AR  # noqa: E402
import _shapes as S  # noqa: E402

HERE = Path(__file__).resolve().parent
STYLES_DIR = HERE.parent / "styles"
A = "{%s}" % S.A_NS
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
P = "{%s}" % P_NS

ROLES = ("card", "callout", "keyfigure", "label", "panel", "decoration")
NEVER = ("logo", "frame")  # never restyled, whatever the style says
TAG = re.compile(r"^(card|callout|keyfigure|label|panel|frame|decoration)#(\d+)(?: > (.+))?$")
TEXT_TYPES = ("text", "bullets")
TOL = int(0.02 * S.EMU_PER_INCH)  # a field "inside" a container may overhang it by this much
DEFAULT_SCOPE = {
    "layouts": {"density": ["Low", "Medium"], "exclude_categories": ["Table", "Frame"]},
    "visuals": {"kinds": ["diagram", "timeline"]},
}
# round silhouettes: the text they hold is centred as a group (top-left in a pebble
# reads as a mistake); rectangles, tickets, torn and hand-drawn cards keep their layout
CENTRED_KINDS = ("blob", "pebble", "drop", "circle", "cloud", "ellipse")
HOOKS = ("arrangement",)  # reported as not applied unless _arrange lays slides out


class RestyleError(Exception):
    pass


# ------------------------------------------------------------------ style
def load_style(which: str) -> dict:
    """A style by name (styles/<name>.json) or path, validated, with defaults filled in."""
    path = Path(which)
    if not which.endswith(".json"):
        path = STYLES_DIR / f"{which}.json"
    if not path.is_file():
        known = ", ".join(sorted(p.stem for p in STYLES_DIR.glob("*.json")))
        raise RestyleError(f"unknown style '{which}' (known: {known}, or a .json file)")
    style = json.loads(path.read_text(encoding="utf-8"))
    style.setdefault("name", path.stem)
    style.setdefault("seed", 0)
    style.setdefault("keep", 0.6)
    style.setdefault("max_chars", 200)
    style.setdefault("min_chars_per_line", 10)
    style.setdefault("min_cover", 0.85)
    style.setdefault("padding", 0.75)
    scope = json.loads(json.dumps(DEFAULT_SCOPE))
    for lib, conf in (style.get("scope") or {}).items():
        scope.setdefault(lib, {}).update(conf)
    style["scope"] = scope
    roles = style.get("roles") or {}
    for role, conf in roles.items():
        if role in NEVER:
            raise RestyleError(f"style '{style['name']}': the {role} role is never restyled")
        if role not in ROLES:
            raise RestyleError(f"style '{style['name']}': unknown role '{role}' (roles: {', '.join(ROLES)})")
        if conf is None:
            continue
        for cand in conf.get("shapes", []):
            if cand.get("kind") not in S.KINDS:
                raise RestyleError(f"style '{style['name']}', {role}: unknown shape '{cand.get('kind')}'")
    style["roles"] = roles
    try:
        AR.style_arrangements(style)
        AR.style_intensity(style)
    except ValueError as exc:
        raise RestyleError(str(exc)) from None
    try:
        style["ambience"] = AMB.normalise(style.get("ambience"))
    except AMB.AmbienceError as exc:
        raise RestyleError(f"style '{style['name']}': {exc}") from None
    return style


def _seed(style: dict, row_id: str, name: str) -> int:
    """Deterministic seed per container: the style's base seed plus a stable hash of the
    layout id and the container name (so card#1 and card#2 differ, and runs agree)."""
    return int(style["seed"]) + zlib.crc32(f"{row_id}|{name}".encode("utf-8")) % 9973


# ------------------------------------------------------------------ index
def read_index(path: Path) -> Tuple[List[str], List[dict]]:
    raw = path.read_bytes().decode("utf-8-sig")
    rows = list(csv.reader(io.StringIO(raw, newline=""), delimiter=";"))
    header = rows[0]
    return header, [dict(zip(header, r)) for r in rows[1:]]


def write_index(path: Path, header: List[str], rows: List[dict]) -> None:
    buf = io.StringIO(newline="")
    w = csv.writer(buf, delimiter=";", quoting=csv.QUOTE_MINIMAL)
    w.writerow(header)
    for r in rows:
        w.writerow([r.get(h, "") for h in header])
    path.write_bytes(buf.getvalue().encode("utf-8"))


def parse_fields(raw: str) -> Dict[str, dict]:
    out = {}
    for part in (raw or "").split(";"):
        part = part.strip()
        if not part:
            continue
        bits = part.split("|")
        flags = [f.strip() for f in (bits[3] if len(bits) > 3 else "").split(",") if f.strip()]
        limit = int(bits[2]) if len(bits) > 2 and bits[2].strip().isdigit() else None
        role = next((f.split("=", 1)[1] for f in flags if f.startswith("role=")), None)
        out[bits[0]] = {"type": bits[1] if len(bits) > 1 else "", "limit": limit, "role": role}
    return out


def set_limits(raw: str, limits: Dict[str, int]) -> str:
    """The fields column with new limits for some fields, everything else byte-identical."""
    parts = raw.split("; ")
    for i, part in enumerate(parts):
        bits = part.split("|")
        if bits[0] in limits and len(bits) > 2:
            bits[2] = str(limits[bits[0]])
            parts[i] = "|".join(bits)
    return "; ".join(parts)


def in_scope(row: dict, lib: str, style: dict) -> Optional[str]:
    """None when the row is in scope, else the reason it is left alone."""
    conf = style["scope"].get(lib, {})
    if lib == "visuals":
        kinds = conf.get("kinds") or []
        return None if row.get("kind") in kinds else f"visual kind '{row.get('kind')}' out of scope"
    if row.get("category") in (conf.get("exclude_categories") or []) or row.get("category") in ("Table", "Frame"):
        return f"category '{row.get('category')}' out of scope"
    if row.get("density") not in (conf.get("density") or []):
        return f"density '{row.get('density')}' out of scope"
    return None


# ------------------------------------------------------------------ shapes
def box(sh) -> Tuple[int, int, int, int]:
    return int(sh.left or 0), int(sh.top or 0), int(sh.width or 0), int(sh.height or 0)


def inside(inner, outer, tol: int = TOL) -> bool:
    x, y, w, h = inner
    X, Y, W, H = outer
    return x >= X - tol and y >= Y - tol and x + w <= X + W + tol and y + h <= Y + H + tol


def collections(shapes):
    """Every shape collection of a slide (the slide, then each group), as lists."""
    out = [list(shapes)]
    for sh in shapes:
        if sh.shape_type == 6:  # group
            out.extend(collections(sh.shapes))
    return out


def prst(sh) -> Optional[str]:
    if sh._element.tag != P + "sp":
        return None
    g = sh._element.spPr.find(A + "prstGeom")
    return g.get("prst") if g is not None else None


def font_info(sh, char_em: dict) -> Tuple[float, float, bool]:
    """(size pt, char_em, caps) of a text shape: the largest run size, its font role."""
    el = sh._element
    sizes = [int(r.get("sz")) for r in el.iter(A + "rPr", A + "endParaRPr", A + "defRPr") if r.get("sz")]
    size = max(sizes) / 100 if sizes else 18.0
    caps = any(r.get("cap") == "all" for r in el.iter(A + "rPr"))
    major = any((lt.get("typeface") or "").startswith("+mj") for lt in el.iter(A + "latin"))
    return size, float(char_em["heading" if major else "body"]), caps


def body_insets(sh) -> Tuple[int, int, int, int]:
    bp = sh._element.find(".//" + A + "bodyPr")
    d = {"lIns": 91440, "tIns": 45720, "rIns": 91440, "bIns": 45720}
    if bp is None:
        return tuple(d.values())  # type: ignore[return-value]
    return tuple(int(bp.get(k, v)) for k, v in d.items())  # type: ignore[return-value]


def paragraphs(sh) -> List[str]:
    if not sh.has_text_frame:
        return []
    return [p.text for p in sh.text_frame.paragraphs]


def lines_needed(paras: List[str], width_pt: float, unit: float) -> Tuple[int, Optional[str]]:
    """Greedy wrap, paragraph by paragraph: (lines, word that would break or None)."""
    total = 0
    for para in paras:
        words = para.split()
        if not words:
            total += 1
            continue
        line, used = "", 1
        for wd in words:
            if len(wd) * unit * S.WORD_SAFETY > width_pt:
                return total + used, wd
            cand = f"{line} {wd}" if line else wd
            if len(cand) * unit <= width_pt:
                line = cand
            else:
                used, line = used + 1, wd
        total += used
    return total, None


def check_field(sh, new_box, old_limit: Optional[int], text_paras: List[str], style: dict, char_em: dict,
                keep: float, max_words: Optional[int] = None) -> dict:
    """Can text field ``sh`` live in ``new_box``? Returns {ok, reason, limit, words}."""
    size, em, caps = font_info(sh, char_em)
    unit = size * em * (1.2 if caps else 1.0)
    li, ti, ri, bi = body_insets(sh)
    w_emu, h_emu = new_box[2] - li - ri, new_box[3] - ti - bi
    w_pt, h_pt = w_emu / S.EMU_PER_PT, h_emu / S.EMU_PER_PT
    res = {"ok": False, "size_pt": size}
    if w_pt <= 0 or h_pt <= 0:
        res["reason"] = "no room left for the text box"
        return res
    per_line = math.floor(w_pt / unit)
    lines = math.floor(h_pt / (size * 1.2))
    if lines < 1:
        res["reason"] = f"not one line of {size:g} pt fits"
        return res
    formula = max(4, math.floor(per_line * lines * 0.85))
    limit = min(old_limit, formula) if old_limit is not None else formula
    res["limit"] = limit
    if old_limit is not None and old_limit > style["min_chars_per_line"] and per_line < style["min_chars_per_line"]:
        res["reason"] = f"{per_line} characters per line, words would break"
        return res
    if old_limit is not None and limit < old_limit:
        need = max(4, math.ceil(keep * old_limit))
        if limit < need:
            res["reason"] = f"limit {old_limit} -> {limit}, below {need} ({keep:.0%} kept)"
            return res
    used, broken = lines_needed(text_paras, w_pt, unit)
    if broken:
        res["reason"] = f"the word '{broken}' would break"
        return res
    if used > lines:
        res["reason"] = f"text needs {used} lines, the box holds {lines}"
        return res
    words = S._wrap_count(["abcde"] * 400, w_emu, h_emu, size, em, caps)
    words = min(words, (limit + 1) // 6) if limit >= 6 else min(words, 1)
    if max_words is not None:
        words = min(words, max_words)
        n = sum(len(p.split()) for p in text_paras)
        if n > max_words:
            res["reason"] = f"{n} words, the shape holds {max_words} at most"
            return res
    res.update(ok=True, words=max(words, 1))
    return res


def fit_aspect(kind: str, b, params: dict, min_cover: float) -> Optional[Tuple[int, int, int, int]]:
    """The container box shrunk (centred) to the shape's aspect bounds, or None when it
    would lose more than ``min_cover`` of a side."""
    x, y, w, h = b
    lo, hi = S.aspect_bounds(kind, **params)
    a = w / h
    if a > hi:
        nw = int(h * hi)
        if nw < w * min_cover:
            return None
        return x + (w - nw) // 2, y, nw, h
    if a < lo:
        nh = int(w / lo)
        if nh < h * min_cover:
            return None
        return x, y + (h - nh) // 2, w, nh
    return b


def map_boxes(boxes: List[Tuple[int, int, int, int]], safe, min_h=None) -> List[Tuple[int, int, int, int]]:
    """Move the field boxes into the safe box: untouched when they already sit inside,
    else shifted the least possible and, when larger, narrowed around their union.
    Too tall: a vertical stack gives up its gaps and the slack under each text first
    (``min_h(i, new_width)`` is the least height box i needs), other sets scale down."""
    ux0 = min(b[0] for b in boxes)
    uy0 = min(b[1] for b in boxes)
    ux1 = max(b[0] + b[2] for b in boxes)
    uy1 = max(b[1] + b[3] for b in boxes)
    U = (ux0, uy0, ux1 - ux0, uy1 - uy0)
    if inside(U, safe, 0):
        return boxes
    sx = min(1.0, safe[2] / U[2])
    nx = min(max(U[0], safe[0]), safe[0] + safe[2] - U[2] * sx)
    xs = [(int(round(nx + (x - U[0]) * sx)), int(round(w * sx))) for x, _, w, _ in boxes]
    if U[3] <= safe[3]:
        ny = min(max(U[1], safe[1]), safe[1] + safe[3] - U[3])
        return [(xs[i][0], int(round(ny + y - U[1])), xs[i][1], h) for i, (_, y, _, h) in enumerate(boxes)]
    order = sorted(range(len(boxes)), key=lambda i: (boxes[i][1], boxes[i][0]))
    stacked = all(boxes[b][1] >= boxes[a][1] + boxes[a][3] - TOL for a, b in zip(order, order[1:]))
    if stacked and min_h is not None:
        mins = [min(boxes[i][3], max(0, min_h(i, xs[i][1]))) for i in range(len(boxes))]
        flex = U[3] - sum(mins)
        t = (safe[3] - sum(mins)) / flex if flex > 0 else -1
        if t >= 0:
            out: List[Optional[Tuple[int, int, int, int]]] = [None] * len(boxes)
            y, prev_bottom = float(safe[1]), None
            for i in order:
                if prev_bottom is not None:
                    y += max(0, boxes[i][1] - prev_bottom) * t
                h = mins[i] + (boxes[i][3] - mins[i]) * t
                out[i] = (xs[i][0], int(round(y)), xs[i][1], int(round(h)))
                y += h
                prev_bottom = boxes[i][1] + boxes[i][3]
            return out  # type: ignore[return-value]
    sy = safe[3] / U[3]
    return [(xs[i][0], int(round(safe[1] + (y - U[1]) * sy)), xs[i][1], int(round(h * sy)))
            for i, (_, y, _, h) in enumerate(boxes)]


def min_height(sh, width: int, paras: List[str], char_em: dict) -> int:
    """Least height (EMU) a text shape needs at ``width`` for its text, one line at least."""
    size, em, caps = font_info(sh, char_em)
    unit = size * em * (1.2 if caps else 1.0)
    li, ti, ri, bi = body_insets(sh)
    n, _ = lines_needed(paras or [""], max(1, width - li - ri) / S.EMU_PER_PT, unit)
    return int(math.ceil(max(1, n) * size * 1.2 * S.EMU_PER_PT)) + ti + bi + 1


def keep_padding(safe, b, cb, boxes, share: float):
    """The safe box, pulled in so the text keeps ``share`` of the room it had from the
    container's edges (a label must not touch the edge of a ticket)."""
    if not boxes:
        return safe
    ux0 = min(x for x, _, _, _ in boxes)
    uy0 = min(y for _, y, _, _ in boxes)
    ux1 = max(x + w for x, _, w, _ in boxes)
    uy1 = max(y + h for _, y, _, h in boxes)
    pad = max(0, min(ux0 - cb[0], uy0 - cb[1], cb[0] + cb[2] - ux1, cb[1] + cb[3] - uy1)) * share
    x0, y0 = max(safe[0], b[0] + pad), max(safe[1], b[1] + pad)
    x1 = min(safe[0] + safe[2], b[0] + b[2] - pad)
    y1 = min(safe[1] + safe[3], b[1] + b[3] - pad)
    if x1 <= x0 or y1 <= y0:
        return None
    return int(x0), int(y0), int(x1 - x0), int(y1 - y0)


def torn_edges(b, slide_wh) -> str:
    """Edges of a panel that face the slide (the others bleed off it)."""
    x, y, w, h = b
    W, H = slide_wh
    sides = []
    if y > TOL:
        sides.append("top")
    if x + w < W - TOL:
        sides.append("right")
    if y + h < H - TOL:
        sides.append("bottom")
    if x > TOL:
        sides.append("left")
    if len(sides) == 4 or not sides:
        return "bottom"
    return "+".join(sides)


def ensure_stroke(sh, width_pt: float, colour: Optional[str] = None) -> None:
    """Shapes with outline-only strokes (hand-drawn) need a line: a theme colour
    (``colour``, e.g. "tx2") or, by default, the fill colour."""
    sppr = sh._element.spPr
    ln = sppr.find(A + "ln")
    if ln is not None and len(ln):
        return
    fill = sppr.find(A + "solidFill")
    if fill is None:
        return
    if ln is None:
        ln = etree.SubElement(sppr, A + "ln")
    ln.set("w", str(int(width_pt * S.EMU_PER_PT)))
    if colour:
        etree.SubElement(etree.SubElement(ln, A + "solidFill"), A + "schemeClr", val=colour)
    else:
        ln.append(copy.deepcopy(fill))


def set_box(sh, b) -> None:
    sh.left, sh.top, sh.width, sh.height = (int(v) for v in b)


def square_box(b, min_cover: float):
    """The box shrunk (centred) towards a square, never below ``min_cover`` of its longer side."""
    x, y, w, h = b
    if h > w:
        nh = max(w, math.ceil(h * min_cover))
        return x, y + (h - nh) // 2, w, nh
    if w > h:
        nw = max(h, math.ceil(w * min_cover))
        return x + (w - nw) // 2, y, nw, h
    return b


def centre_boxes(boxes, safe):
    """The boxes moved together so their union sits in the middle of the safe box."""
    if not boxes:
        return boxes
    ux0 = min(b[0] for b in boxes)
    uy0 = min(b[1] for b in boxes)
    ux1 = max(b[0] + b[2] for b in boxes)
    uy1 = max(b[1] + b[3] for b in boxes)
    dx = round(safe[0] + safe[2] / 2 - (ux0 + ux1) / 2)
    dy = round(safe[1] + safe[3] / 2 - (uy0 + uy1) / 2)
    return [(x + dx, y + dy, w, h) for x, y, w, h in boxes]


def compact(boxes, held, results, check):
    """Stacked fields in a round shape: each text box loses the height its character
    limit does not need (same limit, same words), and the stack closes up keeping its
    gaps, so the group can sit in the middle of the shape. Ornaments keep their size."""
    out = list(boxes)
    for i, sh in enumerate(held):
        x, y, w, h = out[i]
        size = font_info(sh, {"heading": 0.52, "body": 0.52})[0]
        step = int(size * 1.2 * S.EMU_PER_PT)
        want = results[sh.name]
        nh = h
        while nh - step > 0:  # drop one line at a time while the limit and words hold
            r = check(sh, (x, y, w, nh - step))
            if not r["ok"] or r["limit"] != want["limit"] or r["words"] != want["words"]:
                break
            nh -= step
        out[i] = (x, y, w, nh)
    order = sorted(range(len(boxes)), key=lambda i: (boxes[i][1], boxes[i][0]))
    stacked = all(boxes[b][1] >= boxes[a][1] + boxes[a][3] - TOL for a, b in zip(order, order[1:]))
    if not stacked:
        return out
    y = boxes[order[0]][1]
    for k, i in enumerate(order):
        if k:
            a = order[k - 1]
            y += max(0, boxes[i][1] - (boxes[a][1] + boxes[a][3]))
        out[i] = (out[i][0], y, out[i][2], out[i][3])
        y += out[i][3]
    return out


def anchor_middle(sh) -> None:
    bp = sh._element.find(".//" + A + "bodyPr")
    if bp is not None:
        bp.set("anchor", "ctr")


def center_text(sh) -> None:
    for p in sh._element.iter(A + "p"):
        ppr = p.find(A + "pPr")
        if ppr is None:
            ppr = etree.Element(A + "pPr")
            p.insert(0, ppr)
        ppr.set("algn", "ctr")


# ------------------------------------------------------------------ core
class Ctx:
    def __init__(self, style: dict, char_em: dict, slide_wh, deck: bool):
        self.style, self.char_em, self.slide_wh, self.deck = style, char_em, slide_wh, deck


def restyle_container(ctx: Ctx, cont, role: str, row_id: str, coll, text_fields: Dict[str, Optional[int]],
                      containers: set, deco: set) -> dict:
    """Try each candidate shape of the role on one container. Returns a report item."""
    style = ctx.style
    conf = style["roles"][role]
    cb = box(cont)
    m = TAG.match(cont.name)
    linked = m.group(3) if m else cont.name
    self_text = cont.name in text_fields
    held = [sh for sh in coll if sh is not cont and sh.name in text_fields and sh.name not in containers
            and sh.has_text_frame and inside(box(sh), cb)]
    nested = [sh for sh in coll if sh is not cont and sh.name in containers and inside(box(sh), cb)
              and box(sh) != cb]
    # ornaments inside the container (a rule under a number) move with the text
    ornaments = [sh for sh in coll if sh is not cont and sh not in held and sh.name not in containers
                 and (sh.name.startswith("decoration#") or (sh.name in deco and sh.name not in text_fields))
                 and inside(box(sh), cb) and box(sh) != cb]
    item = {"shape": cont.name, "role": role}
    texts = {sh.name: paragraphs(sh) for sh in held + ([cont] if self_text else [])}
    budget = (sum(len(" ".join(t)) for t in texts.values()) if ctx.deck
              else sum(text_fields[n] or 0 for n in texts))
    cap = conf.get("max_chars", style["max_chars"])
    if budget > cap:
        item.update(status="text-heavy", reason=f"{budget} characters, more than {cap}")
        return item
    linked_sh = next((sh for sh in held if sh.name == linked), cont if self_text else (held[0] if held else None))
    size = font_info(linked_sh, ctx.char_em)[0] if linked_sh is not None else S.REFERENCE_PT
    tried = []
    for cand in conf.get("shapes", []):
        kind = cand["kind"]
        params = dict(cand.get("params") or {})
        if params.get("seed") == "auto":
            params["seed"] = _seed(style, row_id, cont.name)
        if kind == "torn" and params.get("edges") == "auto":
            params["edges"] = torn_edges(cb, ctx.slide_wh)
        if kind == "handdrawn" and params.get("form") == "auto":
            params["form"] = "diamond" if prst(cont) == "diamond" else "rect"
            if params["form"] == "diamond":
                params.setdefault("pad", 0.0)  # the text box of a diamond is already its middle half
        why = _try(ctx, cont, kind, params, cand, cb, held, nested, ornaments, texts, text_fields, self_text, size,
                   linked)
        if isinstance(why, dict):
            item.update(status="restyled", kind=kind, params=params, fields=why)
            if tried:
                item["tried"] = tried
            return item
        tried.append({"kind": kind, "reason": why})
    item.update(status="fallback", reason="no candidate shape fits", tried=tried)
    return item


def _try(ctx, cont, kind, params, cand, cb, held, nested, ornaments, texts, text_fields, self_text, size, linked):
    """Apply one candidate if every check passes; return field results or the reason."""
    style = ctx.style
    try:
        b = cb
        margin = float(cand.get("margin", 0))
        if margin:  # breathing room: a pebble that bulges to its box touches its neighbours
            mm = int(margin * min(cb[2], cb[3]))
            b = (cb[0] + mm, cb[1] + mm, cb[2] - 2 * mm, cb[3] - 2 * mm)
        if cand.get("square"):  # a key figure reads as a disc, not a tall oval
            b = square_box(b, style["min_cover"])
        free = bool(cand.get("any_aspect"))  # e.g. a torn edge along a band of any length
        if not free:
            b = fit_aspect(kind, b, params, style["min_cover"])
            if b is None:
                return "box proportions too far from the shape's"
            if not S.fits(kind, b[2], b[3], size_pt=size, **params):
                return f"shape smaller than its minimum size at {size:g} pt"
        res = S.make_shape(kind, b, strict=not free, **params)
    except S.ShapeError as exc:
        return str(exc)
    safe = keep_padding(res.insets.rect(*b), b, cb, [box(sh) for sh in held], style["padding"])
    if safe is None:
        return "no room left inside the shape"
    for sh in nested:
        if not inside(box(sh), safe):
            return f"nested container '{sh.name}' would stick out"
    keep = float(cand.get("keep", style["keep"]))
    max_words = S.KINDS[kind].max_words
    movers = held + ornaments

    def need(i, width):
        sh = movers[i]
        return min_height(sh, width, texts[sh.name], ctx.char_em) if i < len(held) else box(sh)[3]

    new_boxes = map_boxes([box(sh) for sh in movers], safe, need) if held else []
    results = {}
    plan = list(zip(held, new_boxes))
    if self_text:
        plan.append((cont, safe))

    def check(sh, nb):
        return check_field(sh, nb, None if ctx.deck else text_fields[sh.name], texts[sh.name], style,
                           ctx.char_em, keep, max_words if sh.name == linked else None)

    for sh, nb in plan:
        r = check(sh, nb)
        if not r["ok"]:
            return f"{sh.name}: {r['reason']}"
        results[sh.name] = {"limit": r["limit"], "words": r["words"]}
        if not ctx.deck:
            results[sh.name]["old_limit"] = text_fields[sh.name]
    centred = kind in CENTRED_KINDS
    if centred and new_boxes:
        if len(held) > 1 and not ctx.deck:
            new_boxes = compact(new_boxes, held, results, check)
        new_boxes = centre_boxes(new_boxes, safe)
    # every check passed: apply
    if b != cb:
        set_box(cont, b)
    S.set_geometry(cont, res.geom)
    if any(not p.fill for p in S.outline(kind, b[2], b[3], strict=False, **params)):
        ensure_stroke(cont, float(cand.get("stroke_pt", 1.25)), cand.get("stroke"))
    for sh, nb in zip(movers, new_boxes):
        if nb != box(sh):
            set_box(sh, nb)
    if cand.get("align") == "center":
        for sh in held + ([cont] if self_text else []):
            center_text(sh)
    if centred:
        # one text in a round shape: middle of the shape, whatever the text's length
        if self_text and not held:
            anchor_middle(cont)
            center_text(cont)
        elif len(held) == 1 and not self_text:
            anchor_middle(held[0])
    return results


def restyle_slide(ctx: Ctx, slide, row_id: str, fields: Dict[str, dict], scope_reason: Optional[str]) -> List[dict]:
    """Every container of one slide (or visual). ``fields`` is the index entry
    (library), or {} for a deck, where any text shape counts as a field."""
    items = []
    for coll in collections(slide.shapes):
        if fields:
            text_fields = {n: f["limit"] for n, f in fields.items() if f["type"] in TEXT_TYPES}
            flagged = {n: f["role"] for n, f in fields.items() if f["role"]}
            deco = {n for n, f in fields.items() if f["type"] == "deco"}
        else:
            deco = set()
            text_fields = {sh.name: None for sh in coll if sh.has_text_frame and not TAG.match(sh.name)
                           and sh.text_frame.text.strip()}
            flagged = {}
        cands = []
        for sh in coll:
            m = TAG.match(sh.name)
            role = m.group(1) if m else flagged.get(sh.name)
            if role:
                cands.append((sh, role))
        names = {sh.name for sh, role in cands if role not in NEVER}
        for sh, role in cands:
            item = {"id": row_id, "shape": sh.name, "role": role}
            if role in NEVER:
                item.update(status="kept", reason=f"the {role} role is never restyled")
            elif not ctx.style["roles"].get(role):
                item.update(status="kept", reason=f"the style leaves {role} alone")
            elif scope_reason:
                item.update(status="out-of-scope", reason=scope_reason)
            elif prst(sh) is None or prst(sh) == "line" or not sh.width or not sh.height:
                item.update(status="kept", reason="not an area shape")
            else:
                item = {"id": row_id, **restyle_container(ctx, sh, role, row_id, coll, text_fields, names, deco)}
            items.append(item)
    return items


# ------------------------------------------------------------------ files
def save_deterministic(prs, path: Path) -> None:
    """Save with fixed zip timestamps, so the same input gives the same bytes."""
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    out = io.BytesIO()
    with zipfile.ZipFile(buf) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zi = zipfile.ZipInfo(info.filename, date_time=(1980, 1, 1, 0, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o600 << 16
            zout.writestr(zi, zin.read(info.filename))
    path.write_bytes(out.getvalue())


def _char_em(folder: Path) -> dict:
    tokens = folder / "style.json"
    em = {"heading": 0.52, "body": 0.52}
    if tokens.is_file():
        em.update(json.loads(tokens.read_text(encoding="utf-8")).get("char_em") or {})
    return em


def _hooks(style: dict) -> dict:
    return {h: {"requested": style[h], "applied": False, "note": "later layer, not applied by this version"}
            for h in HOOKS if style.get(h)}


def summarise(items: List[dict]) -> dict:
    out: Dict[str, dict] = {}
    for it in items:
        s = out.setdefault(it["role"], {})
        s[it["status"]] = s.get(it["status"], 0) + 1
    total = {}
    for s in out.values():
        for k, v in s.items():
            total[k] = total.get(k, 0) + v
    by_role = {r: dict(sorted(s.items())) for r, s in sorted(out.items())}
    return {"total": dict(sorted(total.items())), "by_role": by_role}


def restyle_library(src: Path, style: dict, out: Path) -> dict:
    """Restyle both libraries of a folder into a new style pack folder."""
    char_em = _char_em(src)
    out.mkdir(parents=True, exist_ok=True)
    report = {"style": style["name"], "input": src.name, "kind": "library", "char_em": char_em, "libraries": {}}
    all_items = []
    for lib in ("layouts", "visuals"):
        pptx, idx = src / f"{lib}.pptx", src / f"{lib}.csv"
        if not pptx.is_file() or not idx.is_file():
            continue
        header, rows = read_index(idx)
        prs = Presentation(str(pptx))
        ctx = Ctx(style, char_em, (prs.slide_width, prs.slide_height), deck=False)
        variants, items = AR.add_variants(sys.modules[__name__], ctx, prs, rows, lib)  # units first (#1607)
        rows.extend(variants)
        by_slide = {int(r["slide"]): r for r in rows if r.get("slide", "").isdigit()}
        for n, slide in enumerate(prs.slides, 1):
            row = by_slide.get(n)
            if row is None:
                continue
            fields = parse_fields(row.get("fields", ""))
            scope = in_scope(row, lib, style)
            got = restyle_slide(ctx, slide, row["id"], fields, scope)
            limits = {}
            for it in got:
                for name, f in (it.get("fields") or {}).items():
                    if f["limit"] != f["old_limit"]:
                        limits[name] = f["limit"]
            if limits:
                row["fields"] = set_limits(row["fields"], limits)
            items.extend(got)
        amb = AMB.apply_library(prs, lib, by_slide, style)
        save_deterministic(prs, out / f"{lib}.pptx")
        write_index(out / f"{lib}.csv", header, rows)
        report["libraries"][lib] = {"summary": summarise(items), "items": items}
        if amb:
            report["libraries"][lib]["ambience"] = amb
        all_items.extend(items)
    if not report["libraries"]:
        raise RestyleError(f"{src}: no layouts.pptx/.csv or visuals.pptx/.csv found")
    for extra in ("theme-contract.json", "style.json"):
        if (src / extra).is_file():
            shutil.copyfile(src / extra, out / extra)
    (out / "restyle.json").write_text(json.dumps(style, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    report["summary"] = summarise(all_items)
    report["hooks"] = _hooks(style)
    if AR.style_arrangements(style):
        report["hooks"]["arrangement"] = {"requested": AR.style_arrangements(style), "applied": True}
    (out / "restyle-report.json").write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    return report


def restyle_deck(src: Path, style: dict, out: Path) -> dict:
    """Restyle a built deck. Roles come from the shape names; slides with a chart or a
    table are left alone; the fit check uses the deck's real text."""
    prs = Presentation(str(src))
    ctx = Ctx(style, {"heading": 0.52, "body": 0.52}, (prs.slide_width, prs.slide_height), deck=True)
    slides = []
    for n, slide in enumerate(prs.slides, 1):
        data = any(getattr(sh, "has_chart", False) or getattr(sh, "has_table", False)
                   for coll in collections(slide.shapes) for sh in coll)
        slides.append((f"slide {n}", slide, "slide with a chart or a table" if data else None))
    items, arranged = AR.arrange_deck(sys.modules[__name__], ctx, slides)  # units first (#1607)
    for key, slide, reason in slides:
        items.extend(restyle_slide(ctx, slide, key, {}, reason))
    amb = AMB.apply_deck(prs, style)
    save_deterministic(prs, out)
    hooks = _hooks(style)
    if arranged:
        hooks["arrangement"] = arranged
    report = {"style": style["name"], "input": src.name, "kind": "deck", "summary": summarise(items),
              "hooks": hooks, "items": items}
    if amb:
        report["ambience"] = amb
    return report


def run(src: Path, style: dict, out: Optional[Path] = None, force: bool = False) -> Tuple[Path, dict]:
    src = src.resolve()
    deck = src.suffix.lower() == ".pptx"
    if out is None:
        name = f"{src.stem}-{style['name']}.pptx" if deck else f"{src.name}-{style['name']}"
        out = src.with_name(name)
    out = out.resolve()
    if out == src or (not deck and (out == src or src in out.parents or out in src.parents)):
        raise RestyleError("the output must be a new file or folder, never the input")
    if out.exists() and not force:
        raise RestyleError(f"{out} already exists (use --force to replace it)")
    if deck:
        out.parent.mkdir(parents=True, exist_ok=True)
        report = restyle_deck(src, style, out)
    else:
        if not src.is_dir():
            raise RestyleError(f"{src}: not a library folder nor a .pptx deck")
        if out.exists():
            shutil.rmtree(out)
        report = restyle_library(src, style, out)
    return out, report


def override(style: dict, arrangement: Optional[str] = None, intensity: Optional[str] = None,
             slides: Optional[str] = None) -> dict:
    """The style with the arrangement, intensity or chosen slides of this run (validated)."""
    style = dict(style)
    if arrangement:
        style["arrangement"] = [a.strip() for a in arrangement.split(",") if a.strip()]
    if intensity:
        style["intensity"] = intensity
    if slides:
        try:
            chosen = [int(n) for n in slides.split(",") if n.strip()]
        except ValueError:
            raise RestyleError(f"--slides takes slide numbers, not '{slides}'") from None
        style["dosage"] = {**(style.get("dosage") or {}), "slides": chosen}
    try:
        AR.style_arrangements(style)
        AR.style_intensity(style)
    except ValueError as exc:
        raise RestyleError(str(exc)) from None
    return style


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("input", help="library folder or built deck (.pptx)")
    ap.add_argument("--style", required=True, help="style name (styles/) or .json file")
    ap.add_argument("--out", help="output folder (library) or .pptx (deck)")
    ap.add_argument("--report", help="also write the JSON report here")
    ap.add_argument("--force", action="store_true", help="replace an existing output")
    ap.add_argument("--arrangement", help="arrangements, comma separated (stagger, cascade, arc, deconstructed)")
    ap.add_argument("--intensity", help="deck: how many slides get arranged (calm, balanced, bold, all)")
    ap.add_argument("--slides", help="deck: arrange these slides only (numbers, comma separated)")
    args = ap.parse_args(argv)
    try:
        style = load_style(args.style)
        if args.arrangement or args.intensity or args.slides:
            style = override(style, args.arrangement, args.intensity, args.slides)
        out, report = run(Path(args.input), style, Path(args.out) if args.out else None, args.force)
    except (RestyleError, OSError, ValueError) as exc:
        print(json.dumps({"error": str(exc)}))
        return 2
    if args.report:
        Path(args.report).write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    elif report["kind"] == "deck":
        out.with_suffix(".report.json").write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"output": str(out), "style": style["name"], "summary": report["summary"]["total"]}, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
