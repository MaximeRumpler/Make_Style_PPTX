"""Arrangement layer: lay out the repeated units of a layout another way than the plain grid.

A layout whose ``arrange`` column in Make_Clean's index says ``row:3``, ``column:5:ordered``…
holds identical units (every field with the same number plus the container that holds them).
This module moves whole units to deterministic positions that depend only on the arrangement,
the kind (row or column) and the number of units:

- ``stagger``        every other unit shifted (down in a row, right in a column)
- ``cascade``        each unit a step further than the previous one (a staircase)
- ``arc``            units follow an arch: the middle ones higher (row) or further right (column)
- ``deconstructed``  the 12-column grid of Make_Clean with exact moves: starts in half-column
                     steps (at most 1.5 columns from the plain grid), drops in baseline units,
                     spans in half columns; a fixed table per unit count

Units are only translated and resized; fonts never change. Inside a unit the fields keep their
order and offsets; when a unit gets shorter, the gaps and the slack under each text give way
first, never below the height that keeps ``keep`` of a field's character limit. The arrangement
is applied only when every check passes, else the slide keeps its plain grid and the report says
why: each text field still fits (same checks as the restyler), no unit overlaps another or a text
box of the slide, every unit stays inside the margins and the content band, and the visual
reading order (left to right in a row, top to bottom in a column) is the order of the units.

Geometry (``positions``, ``fit_stack`` and the checks) is pure and works in EMU; ``arrange``
is the hook called by restyle.py for each slide of a library.
"""
from __future__ import annotations

import math
import re
from typing import Dict, List, Optional, Sequence, Tuple

EMU = 914400  # per inch
Rect = Tuple[int, int, int, int]

# Make_Clean's grid (library/theme/tokens.json): 13.333 x 7.5 in slide, 0.6 in side margins,
# 0.3 in gutter, content band from 2.2 in, 4.55 in high. Twelve columns over the content width.
GRID = {"w": 13.333, "margin_x": 0.6, "gutter": 0.3, "content_top": 2.2, "content_bottom": 6.75,
        "columns": 12, "baseline": 0.15}
NAMES = ("stagger", "cascade", "arc", "deconstructed")
KINDS = ("row", "column")

# Offsets of the plain arrangements: rows in baseline units (vertical), columns in half
# columns (horizontal). Total amplitude, whatever the number of units.
ROW_OFFSET = {"stagger": 6, "cascade": 8, "arc": 6}
COLUMN_OFFSET = {"stagger": 2, "cascade": 3, "arc": 2}

# Deconstructed grid, rows: per unit (start, span) in half columns and drop in baselines.
# Rules (checked by the tests): start within 3 half columns (1.5 columns) of the plain grid,
# one span for every unit (a set stays a set: same width, same shape), never wider than the plain
# one, units side by side never closer than one gutter. The rhythm comes from unequal gaps and drops.
DECONSTRUCTED_ROW = {
    2: [(0, 10, 0), (14, 10, 5)],
    3: [(0, 7, 0), (9, 7, 6), (17, 7, 2)],
    4: [(0, 5, 0), (7, 5, 5), (12, 5, 1), (19, 5, 6)],
}
# Deconstructed grid, columns: indent of each unit in half columns (cycled); every unit loses
# the largest indent used from its width, so the right edges stay on one line.
DECONSTRUCTED_COLUMN = (0, 3, 1, 2, 0, 3)
MAX_SHIFT_HALF = 3

UNIT_NUMBER = re.compile(r"(?<!\S)(\d+)(?!\S)")
ARRANGE = re.compile(r"^(row|column|grid):(\d+)(:ordered)?$")


class ArrangeError(Exception):
    """The arrangement does not apply to these units (the slide keeps its plain grid)."""


def inch(v: float) -> int:
    return int(round(v * EMU))


def pitch() -> float:
    """Column pitch (column + gutter), in inches."""
    g = GRID
    width = g["w"] - 2 * g["margin_x"]
    col = (width - (g["columns"] - 1) * g["gutter"]) / g["columns"]
    return col + g["gutter"]


def half_x(half: int) -> int:
    """Left edge of a start given in half columns from the left margin, EMU."""
    return inch(GRID["margin_x"] + half / 2 * pitch())


def span_w(half: int) -> int:
    """Width of a span given in half columns (gutter excluded), EMU."""
    return inch(half / 2 * pitch() - GRID["gutter"])


def parse_arrange(value: str) -> Optional[Tuple[str, int, bool]]:
    """``row:3`` -> ("row", 3, False); empty or malformed -> None."""
    m = ARRANGE.match((value or "").strip())
    return (m.group(1), int(m.group(2)), bool(m.group(3))) if m else None


def style_arrangement(style: dict) -> Optional[str]:
    """The arrangement a style asks for (``"arc"`` or ``{"name": "arc"}``), validated."""
    spec = style.get("arrangement")
    if not spec:
        return None
    name = spec.get("name") if isinstance(spec, dict) else spec
    if name not in NAMES:
        raise ValueError(f"style '{style.get('name')}': unknown arrangement '{name}' (known: {', '.join(NAMES)})")
    return name


# ------------------------------------------------------------------ positions
def _union(rects: Sequence[Rect]) -> Rect:
    x0 = min(r[0] for r in rects)
    y0 = min(r[1] for r in rects)
    x1 = max(r[0] + r[2] for r in rects)
    y1 = max(r[1] + r[3] for r in rects)
    return x0, y0, x1 - x0, y1 - y0


def _offsets(name: str, n: int, total: float, kind: str = "column") -> List[float]:
    if name == "stagger":
        return [total if i % 2 else 0.0 for i in range(n)]
    if name == "cascade":
        return [total * i / (n - 1) for i in range(n)]
    if kind == "row":  # arc in a row: an arch, the middle units highest, the outer ones lowest
        return [total * (2 * i / (n - 1) - 1) ** 2 for i in range(n)]
    # arc in a column: the middle units bulge to the right
    return [total * (1 - (2 * i / (n - 1) - 1) ** 2) for i in range(n)]


def positions(name: str, kind: str, units: Sequence[Rect], region: Rect) -> List[Rect]:
    """Target rectangle of each unit (same order as ``units``), EMU.

    ``units`` are the units' current boxes; ``region`` is the room they may use (a row may
    grow down to the bottom of the content band). Raises ArrangeError when the arrangement
    does not apply to this kind or count."""
    n = len(units)
    if kind not in KINDS:
        raise ArrangeError(f"{kind} units keep their layout in this version")
    if n < 2:
        raise ArrangeError("fewer than two units")
    if name == "arc" and n < 3:
        raise ArrangeError("an arc needs three units or more")
    rx, ry, rw, rh = region
    bl = inch(GRID["baseline"])
    tol = inch(0.05)
    if kind == "row":
        if name == "deconstructed":
            table = DECONSTRUCTED_ROW.get(n)
            if table is None:
                raise ArrangeError(f"no deconstructed pattern for {n} units in a row")
            if abs(rx - inch(GRID["margin_x"])) > tol or abs(rw - (inch(GRID["w"]) - 2 * inch(GRID["margin_x"]))) > tol:
                raise ArrangeError("the units do not span the 12-column grid")
            drop = max(d for _, _, d in table) * bl
            h = min(max(u[3] for u in units), rh - drop)
            return [(half_x(s), ry + d * bl, span_w(sp), h) for s, sp, d in table]
        total = ROW_OFFSET[name] * bl
        offs = [int(round(o / bl)) * bl for o in _offsets(name, n, total, "row")]
        h = min(max(u[3] for u in units), rh - max(offs))
        return [(u[0], ry + o, u[2], h) for u, o in zip(units, offs)]
    # column: units keep their rows and move sideways, losing the largest indent from their width
    half = inch(pitch() / 2)
    if name == "deconstructed":
        shifts = [DECONSTRUCTED_COLUMN[i % len(DECONSTRUCTED_COLUMN)] * half for i in range(n)]
    else:
        shifts = [int(round(o)) for o in _offsets(name, n, COLUMN_OFFSET[name] * half)]
    loss = max(shifts)
    return [(u[0] + s, u[1], u[2] - loss, u[3]) for u, s in zip(units, shifts)]


# ------------------------------------------------------------------ inside a unit
def fit_stack(boxes: Sequence[Rect], mins: Sequence[int], top: int, height: int) -> Optional[List[Rect]]:
    """Move boxes (one unit's members, already placed horizontally) so they start at ``top``
    and end within ``height``. Boxes that overlap vertically form a band (a mark beside a
    label); bands keep their order, the gaps between them and the slack of each band (its
    height above what its boxes need, ``mins``) shrink by the same share. None when even
    the least heights do not fit."""
    order = sorted(range(len(boxes)), key=lambda i: (boxes[i][1], boxes[i][0]))
    bands: List[List[int]] = []
    bottom = None
    for i in order:
        y, h = boxes[i][1], boxes[i][3]
        if bands and y < bottom:
            bands[-1].append(i)
            bottom = max(bottom, y + h)
        else:
            bands.append([i])
            bottom = y + h
    info = []
    for band in bands:
        y0 = min(boxes[i][1] for i in band)
        y1 = max(boxes[i][1] + boxes[i][3] for i in band)
        need = max(boxes[i][1] - y0 + mins[i] for i in band)
        info.append((y0, y1, min(need, y1 - y0)))
    gaps = [max(0, info[k][0] - info[k - 1][1]) for k in range(1, len(info))]
    total = info[-1][1] - info[0][0]
    out: List[Optional[Rect]] = [None] * len(boxes)
    if total <= height:
        dy = top - info[0][0]
        return [(x, y + dy, w, h) for x, y, w, h in boxes]
    flex = sum(y1 - y0 - need for y0, y1, need in info) + sum(gaps)
    excess = total - height
    if flex <= 0 or excess > flex:
        return None
    t = 1 - excess / flex
    y = float(top)
    for k, (band, (y0, y1, need)) in enumerate(zip(bands, info)):
        if k:
            y += gaps[k - 1] * t
        bh = need + (y1 - y0 - need) * t
        for i in band:
            x, by, w, h = boxes[i]
            off = by - y0
            out[i] = (x, int(round(y + off)), w, int(math.floor(min(h, bh - off))))
        y += bh
    return out  # type: ignore[return-value]


# ------------------------------------------------------------------ checks
def intersects(a: Rect, b: Rect) -> bool:
    """True when the two boxes share some area (touching edges do not count)."""
    return a[0] < b[0] + b[2] and b[0] < a[0] + a[2] and a[1] < b[1] + b[3] and b[1] < a[1] + a[3]


def overlapping(rects: Sequence[Rect]) -> List[Tuple[int, int]]:
    """Pairs of indexes whose boxes overlap."""
    return [(i, j) for i in range(len(rects)) for j in range(i + 1, len(rects)) if intersects(rects[i], rects[j])]


def contains(outer: Rect, inner: Rect, tol: int = 0) -> bool:
    return (inner[0] >= outer[0] - tol and inner[1] >= outer[1] - tol
            and inner[0] + inner[2] <= outer[0] + outer[2] + tol and inner[1] + inner[3] <= outer[1] + outer[3] + tol)


def reading_order(kind: str, rects: Sequence[Rect]) -> List[int]:
    """Indexes of the units in visual reading order: left to right in a row (by centre),
    top to bottom in a column; ties broken the other way."""
    def key(i):
        x, y, w, h = rects[i]
        cx, cy = 2 * x + w, 2 * y + h
        return (cx, cy) if kind == "row" else (cy, cx)
    return sorted(range(len(rects)), key=key)


def bounds(slide_wh: Tuple[int, int], current: Rect) -> Rect:
    """Where units may go: inside the side margins, between the top of the content band (or
    the units' current top, when higher) and its bottom (or their current bottom)."""
    sw = slide_wh[0]
    mx = int(round(sw * GRID["margin_x"] / GRID["w"]))
    top = min(inch(GRID["content_top"]), current[1])
    bottom = max(inch(GRID["content_bottom"]), current[1] + current[3])
    return mx, top, sw - 2 * mx, bottom - top


# ------------------------------------------------------------------ slide hook
def unit_number(name: str) -> Optional[int]:
    nums = UNIT_NUMBER.findall(name)
    return int(nums[0]) if len(nums) == 1 else None


def _expand(r: Rect, d: int) -> Rect:
    return r[0] - d, r[1] - d, r[2] + 2 * d, r[3] + 2 * d


def find_units(R, slide, fields: Dict[str, dict], count: int):
    """({n: [shapes]}, {n: container or None}, separators, others) of a slide, or a reason.
    Members: every field numbered n and every role-tagged shape linked to one; unlinked
    decorations join the one unit whose box (grown by a gutter) holds them, or are
    separators when they sit across two units or more."""
    units: Dict[int, list] = {}
    loose, others = [], []
    for sh in slide.shapes:
        m = R.TAG.match(sh.name)
        link = m.group(3) if m else sh.name
        n = unit_number(link) if link in fields else None
        if m and m.group(1) == "panel" and n:
            return "units held by a panel keep their layout"
        if n is not None and 1 <= n <= count:
            units.setdefault(n, []).append(sh)
        elif m and not m.group(3):
            loose.append(sh)
        else:
            others.append(sh)
    for coll in R.collections(slide.shapes)[1:]:
        for sh in coll:
            if sh.name in fields and unit_number(sh.name):
                return "a unit sits inside a group"
    if sorted(units) != list(range(1, count + 1)):
        return f"units {sorted(units)} found, the index says {count}"
    rects = {n: _union([R.box(sh) for sh in shs]) for n, shs in units.items()}
    g = inch(GRID["gutter"])
    separators = []
    for sh in loose:
        b = R.box(sh)
        homes = [n for n, r in rects.items() if contains(_expand(r, g), b)]
        if len(homes) == 1:
            units[homes[0]].append(sh)
            continue
        hits = [n for n, r in rects.items() if intersects(_expand(r, g // 2 + inch(0.01)), b)]
        if len(homes) > 1 or len(hits) > 1:
            separators.append(sh)
        else:
            others.append(sh)
    containers = {}
    for n, shs in units.items():
        r = _union([R.box(sh) for sh in shs])
        box = next((sh for sh in shs if R.TAG.match(sh.name) and R.TAG.match(sh.name).group(1) != "decoration"
                    and contains(R.box(sh), r, R.TOL)), None)
        containers[n] = box
    return units, containers, separators, others


def _least_height(R, ctx, sh, width: int, old_limit: Optional[int], keep: float) -> int:
    """Least height of a text field at ``width``: its template text, and room for ``keep`` of its limit."""
    paras = R.paragraphs(sh)
    h = R.min_height(sh, width, paras, ctx.char_em)
    if old_limit:
        size, em, caps = R.font_info(sh, ctx.char_em)
        unit = size * em * (1.2 if caps else 1.0)
        li, ti, ri, bi = R.body_insets(sh)
        per_line = math.floor(max(1, width - li - ri) / 12700 / unit)
        if per_line > 0:
            lines = math.ceil(math.ceil(keep * old_limit) / 0.85 / per_line)
            h = max(h, int(math.ceil(lines * size * 1.2 * 12700)) + ti + bi + 1)
    return h


def place_unit(R, ctx, members, container, old: Rect, new: Rect, text_fields: Dict[str, Optional[int]], keep: float):
    """New box of every member of one unit, or a reason. The container (when the unit has
    one) takes the unit's new box; the other members keep their padding inside it."""
    boxes = {sh.name: R.box(sh) for sh in members}
    out: Dict[str, Rect] = {}
    inner_old, inner_new = old, new
    rest = [sh for sh in members if sh is not container]
    if container is not None:
        out[container.name] = new
        if rest:
            u = _union([boxes[sh.name] for sh in rest])
            pl, pt = u[0] - old[0], u[1] - old[1]
            pr, pb = old[0] + old[2] - u[0] - u[2], old[1] + old[3] - u[1] - u[3]
            inner_old = u
            inner_new = (new[0] + pl, new[1] + pt, new[2] - pl - pr, new[3] - pt - pb)
            if inner_new[2] <= 0 or inner_new[3] <= 0:
                return "no room left inside the container"
    if not rest:
        return out
    ox, oy, ow, oh = inner_old
    nx, ny, nw, nh = inner_new
    sx = nw / ow if ow else 1.0
    placed = []
    for sh in rest:
        x, y, w, h = boxes[sh.name]
        # rules and wide boxes follow the unit's width; marks, icons and short labels keep theirs
        nwid = int(round(w * sx)) if (h == 0 or w >= 0.6 * ow) else min(w, nw)
        nxx = min(nx + int(round((x - ox) * sx)), nx + nw - nwid)
        placed.append((nxx, y, nwid, h))
    mins = []
    for sh, b in zip(rest, placed):
        if sh.name in text_fields and sh.has_text_frame:
            mins.append(min(b[3], _least_height(R, ctx, sh, b[2], text_fields[sh.name], keep)))
        else:
            mins.append(b[3])
    stacked = fit_stack(placed, mins, ny, nh)
    if stacked is None:
        return "the unit's texts do not fit its new height"
    for sh, b in zip(rest, stacked):
        out[sh.name] = b
    return out


def arrange(R, ctx, slide, row: dict, fields: Dict[str, dict], lib: str, scope_reason: Optional[str]) -> List[dict]:
    """Hook for restyle.py: arrange the units of one library slide. Returns [] when the style
    or the row asks for nothing, else one report item (arranged, fallback, out of scope or kept).
    The slide is changed only when the arrangement passes every check."""
    name = style_arrangement(ctx.style)
    tag = parse_arrange(row.get("arrange", ""))
    if not name or not tag:
        return []
    kind, count, ordered = tag
    item = {"id": row["id"], "shape": "arrangement", "role": "arrangement", "arrangement": name,
            "arrange": row.get("arrange")}
    if lib != "layouts":
        item.update(status="kept", reason="visuals keep their arrangement in this version")
        return [item]
    if scope_reason:
        item.update(status="out-of-scope", reason=scope_reason)
        return [item]
    why = _arrange(R, ctx, slide, fields, kind, count, name, item)
    if why:
        item.update(status="fallback", reason=f"plain grid kept: {why}")
    else:
        item["status"] = "arranged"
    return [item]


def _arrange(R, ctx, slide, fields, kind, count, name, item) -> Optional[str]:
    found = find_units(R, slide, fields, count)
    if isinstance(found, str):
        return found
    units, containers, separators, others = found
    numbers = sorted(units)
    old = [_union([R.box(sh) for sh in units[n]]) for n in numbers]
    whole = _union(old)
    text_fields = {n: f["limit"] for n, f in fields.items() if f["type"] in R.TEXT_TYPES}
    other_text = [R.box(sh) for sh in others if sh.has_text_frame and sh.text_frame.text.strip()]
    region = whole
    if kind == "row":  # a row may grow down to the content band, short of the next shape below
        bottom = inch(GRID["content_bottom"])
        for b in [R.box(sh) for sh in others]:
            if b[1] >= whole[1] + whole[3] - R.TOL and b[0] < whole[0] + whole[2] and whole[0] < b[0] + b[2]:
                bottom = min(bottom, b[1] - inch(GRID["baseline"]))
        region = (whole[0], whole[1], whole[2], max(whole[3], bottom - whole[1]))
    try:
        targets = positions(name, kind, old, region)
    except ArrangeError as exc:
        return str(exc)
    keep = float(ctx.style["keep"])
    plan: Dict[str, Rect] = {}
    for n, o, t in zip(numbers, old, targets):
        got = place_unit(R, ctx, units[n], containers[n], o, t, text_fields, keep)
        if isinstance(got, str):
            return f"unit {n}: {got}"
        plan.update(got)
    new_units = [_union([plan[sh.name] for sh in units[n]]) for n in numbers]
    # margins and content band
    limit = bounds(ctx.slide_wh, whole)
    for n, r in zip(numbers, new_units):
        if not contains(limit, r, R.TOL):
            return f"unit {n} would leave the margins"
    # units never overlap each other, nor a text box of the slide they did not overlap before
    pairs = overlapping(new_units)
    if pairs:
        a, b = pairs[0]
        return f"units {numbers[a]} and {numbers[b]} would overlap"
    for n, o, r in zip(numbers, old, new_units):
        for b in other_text:
            if intersects(r, b) and not intersects(o, b):
                return f"unit {n} would overlap another text box"
    if reading_order(kind, new_units) != list(range(len(numbers))):
        return "the visual order would not follow the units' order"
    # every text field passes the restyler's fit check in its new box
    results = {}
    for n in numbers:
        for sh in units[n]:
            if sh.name not in text_fields or not sh.has_text_frame:
                continue
            r = R.check_field(sh, plan[sh.name], text_fields[sh.name], R.paragraphs(sh), ctx.style, ctx.char_em, keep)
            if not r["ok"]:
                return f"{sh.name}: {r['reason']}"
            results[sh.name] = {"limit": r["limit"], "old_limit": text_fields[sh.name], "words": r["words"]}
    # apply
    for n in numbers:
        for sh in units[n]:
            if plan[sh.name] != R.box(sh):
                R.set_box(sh, plan[sh.name])
    removed = []
    for sh in separators:
        sh._element.getparent().remove(sh._element)
        removed.append(sh.name)
    item.update(units=len(numbers), fields=results)
    if removed:
        item["removed"] = removed
    return None
