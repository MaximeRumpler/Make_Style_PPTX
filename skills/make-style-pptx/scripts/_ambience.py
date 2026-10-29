"""Ambience layer: a quiet background motif behind the content of low-text slides.

A style opts in with an ``ambience`` key (styles/*.json)::

    "ambience": {"motif": "mist", "density": "low"}

Motifs (all native, editable shapes in theme colours, nothing raster):

- ``mist``: a few soft, overlapping blobs gathered in the emptiest corner;
- ``rain``: thin slanted streaks, kept clear of the text;
- ``grid``: a technical grid of fine lines (or ``"marks": "cross"``) over the emptier side,
  fading out towards the content.

The motif is one group, ``ambience#1 > <motif>``, placed first in the slide's shape
tree (behind everything), marked decorative for screen readers. Colours are theme
references (``tx2`` on light slides, ``bg1`` on dark ones by default) with a
transparency, so a new theme rebrands them. Same slide, style and seed give the same
shapes.

Where it goes: in a library, layouts of the listed categories and densities (covers,
statements, dividers of density Low by default; never tables, frames or visuals).
In a deck, slides without chart, table or picture whose text is short and set large
(``max_chars``, ``max_fields``, ``min_pt``).

Contrast gate: every text run over a motif part is checked against the colour the
part leaves under it (WCAG 2.x, 4.5:1, or 3:1 for large text). A part that would
take a run below its threshold is dropped; when nothing is left, the slide keeps no
motif. Text on its own fill, or over an opaque shape, is not affected.

Optional ``typography``: ``{"major": "<font>", "minor": "<font>"}`` sets the theme
fonts (headings, body). Fonts change text widths: pick a pair close to the original
widths, the character limits are not recomputed.
"""
from __future__ import annotations

import colorsys
import math
import random
import zlib
from typing import Dict, List, Optional, Tuple

from lxml import etree

import _shapes as S

A_NS = S.A_NS
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "{%s}" % A_NS
P = "{%s}" % P_NS
RT_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
DECORATIVE = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}"
ADEC = "http://schemas.microsoft.com/office/drawing/2017/decorative"
EMU = S.EMU_PER_INCH

MOTIFS = ("mist", "rain", "grid")
SCHEME = ("tx1", "tx2", "bg1", "bg2", "dk1", "dk2", "lt1", "lt2",
          "accent1", "accent2", "accent3", "accent4", "accent5", "accent6")
# density tokens: how much of each motif, and how opaque. "low" keeps text contrast
# on the neutral and on most brand themes; the gate catches the rest.
DENSITY = {
    "low": {"mist": {"blobs": 3, "alpha": 0.09}, "rain": {"streaks": 45, "alpha": 0.24, "pt": 0.75},
            "grid": {"pitch": 0.40, "alpha": 0.22, "pt": 0.5}},
    "medium": {"mist": {"blobs": 4, "alpha": 0.14}, "rain": {"streaks": 80, "alpha": 0.32, "pt": 0.75},
               "grid": {"pitch": 0.30, "alpha": 0.30, "pt": 0.5}},
}
TEXT_MIN, LARGE_MIN = 4.5, 3.0
DEFAULTS = {
    "motif": None, "density": "low", "colour": "tx2", "on_dark": "bg1", "seed": 0,
    "categories": ["Cover", "Statement", "Divider"], "layout_density": ["Low"],
    "max_chars": 280, "max_fields": 8, "min_pt": 28, "marks": "lines", "typography": None,
}
NEVER_CATEGORIES = ("Table", "Frame")
GROUP_NAME = "ambience#1"
SOFT = 0.2  # soft edge of a mist blob, share of its shorter side


class AmbienceError(ValueError):
    pass


# ------------------------------------------------------------------ style
def normalise(conf) -> Optional[dict]:
    """The ``ambience`` value of a style, validated, with defaults. None = no ambience."""
    if not conf:
        return None
    if isinstance(conf, str):
        conf = {"motif": conf}
    if not isinstance(conf, dict):
        raise AmbienceError("ambience must be a motif name or an object")
    unknown = set(conf) - set(DEFAULTS)
    if unknown:
        raise AmbienceError(f"ambience: unknown key(s) {', '.join(sorted(unknown))}")
    out = dict(DEFAULTS)
    out.update(conf)
    if out["motif"] is not None and out["motif"] not in MOTIFS:
        raise AmbienceError(f"ambience: unknown motif '{out['motif']}' (motifs: {', '.join(MOTIFS)})")
    if out["density"] not in DENSITY:
        raise AmbienceError(f"ambience: density must be one of {', '.join(DENSITY)}")
    for key in ("colour", "on_dark"):
        if out[key] not in SCHEME:
            raise AmbienceError(f"ambience: {key} must be a theme colour ({', '.join(SCHEME)})")
    if out["marks"] not in ("lines", "cross"):
        raise AmbienceError("ambience: marks must be 'lines' or 'cross'")
    typo = out["typography"]
    if typo is not None:
        if not isinstance(typo, dict) or not set(typo) <= {"major", "minor"} or not all(
                isinstance(v, str) and v.strip() for v in typo.values()):
            raise AmbienceError("ambience: typography is {\"major\": font, \"minor\": font}")
    if out["motif"] is None and typo is None:
        return None
    return out


def slide_seed(conf: dict, style_seed: int, key: str) -> int:
    return int(style_seed) + int(conf["seed"]) + zlib.crc32(f"{key}|{conf['motif']}".encode("utf-8")) % 9973


# ------------------------------------------------------------------ colour
def _hex(h: str):
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def _lin(c: float) -> float:
    return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4


def _gamma(c: float) -> float:
    return c * 12.92 if c <= 0.0031308 else 1.055 * c ** (1 / 2.4) - 0.055


def luminance(rgb) -> float:
    r, g, b = (_lin(c) for c in rgb)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast(fg, bg) -> float:
    hi, lo = sorted((luminance(fg), luminance(bg)), reverse=True)
    return (hi + 0.05) / (lo + 0.05)


def blend(top, alpha: float, below):
    return tuple(t * alpha + b * (1 - alpha) for t, b in zip(top, below))


class Theme:
    """Theme colours and colour map of one slide (master theme, master clrMap,
    layout and slide overrides)."""

    def __init__(self, slide):
        layout = slide.slide_layout
        master = layout.slide_master
        self.colors: Dict[str, Tuple[float, float, float]] = {}
        try:
            root = etree.fromstring(master.part.part_related_by(RT_THEME).blob)
            scheme = root.find(".//" + A + "clrScheme")
            for el in scheme if scheme is not None else []:
                c = el[0] if len(el) else None
                val = None if c is None else (c.get("lastClr") if c.tag == A + "sysClr" else c.get("val"))
                if val and len(val) == 6:
                    self.colors[etree.QName(el).localname] = _hex(val)
        except (KeyError, etree.XMLSyntaxError):
            pass
        self.cmap = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}
        el = master._element.find(P + "clrMap")
        if el is not None:
            self.cmap.update(el.attrib)
        for part in (layout, slide):
            ov = part._element.find(P + "clrMapOvr/" + A + "overrideClrMapping")
            if ov is not None:
                self.cmap.update(ov.attrib)
        self.chain = (slide, layout, master)

    def scheme(self, name: str):
        return self.colors.get(self.cmap.get(name, name), (0.5, 0.5, 0.5))

    def resolve(self, el) -> Optional[Tuple[tuple, float]]:
        """A colour element (srgbClr, schemeClr...) -> (rgb, alpha)."""
        if el is None:
            return None
        tag = etree.QName(el).localname
        if tag == "srgbClr":
            rgb = _hex(el.get("val", "000000"))
        elif tag == "schemeClr":
            rgb = self.scheme(el.get("val", ""))
        elif tag == "sysClr":
            rgb = _hex(el.get("lastClr") or "000000")
        elif tag == "prstClr":
            rgb = (1.0, 1.0, 1.0) if el.get("val") == "white" else (0.0, 0.0, 0.0)
        else:
            return None
        alpha = 1.0
        for m in el:
            name = etree.QName(m).localname
            v = int(m.get("val", "0")) / 100000
            if name in ("lumMod", "lumOff", "satMod"):
                h, lum, s = colorsys.rgb_to_hls(*rgb)
                if name == "lumMod":
                    lum *= v
                elif name == "lumOff":
                    lum += v
                else:
                    s *= v
                rgb = colorsys.hls_to_rgb(h, min(1.0, max(0.0, lum)), min(1.0, max(0.0, s)))
            elif name == "tint":
                rgb = tuple(_gamma(1 - (1 - _lin(c)) * v) for c in rgb)
            elif name == "shade":
                rgb = tuple(_gamma(_lin(c) * v) for c in rgb)
            elif name == "alpha":
                alpha = v
        return tuple(min(1.0, max(0.0, c)) for c in rgb), alpha

    def background(self):
        """The slide background colour (slide, else layout, else master, else bg1)."""
        for part in self.chain:
            bg = part._element.find(P + "cSld/" + P + "bg")
            if bg is None:
                continue
            fill = bg.find(P + "bgPr/" + A + "solidFill")
            if fill is None:
                fill = bg.find(P + "bgRef")
            if fill is not None and len(fill):
                got = self.resolve(fill[0])
                if got:
                    return got[0]
        return self.scheme("bg1")


# ------------------------------------------------------------------ slide reading
def _box(sh) -> Tuple[int, int, int, int]:
    return int(sh.left or 0), int(sh.top or 0), int(sh.width or 0), int(sh.height or 0)


def _flat(shapes) -> list:
    """Every shape in z-order (groups opened)."""
    out = []
    for sh in shapes:
        if sh.shape_type == 6:
            out.extend(_flat(sh.shapes))
        else:
            out.append(sh)
    return out


def _text(sh) -> str:
    """The text of a shape, read from the XML (``sh.text_frame`` would add an empty
    text body to shapes that have none, and the input must stay byte-identical)."""
    return "\n".join("".join(t.text or "" for t in p.iter(A + "t")) for p in sh._element.iter(A + "p"))


def _overlap(a, b) -> int:
    w = min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0])
    h = min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1])
    return max(0, w) * max(0, h)


def _contains(outer, inner) -> bool:
    return (outer[0] <= inner[0] and outer[1] <= inner[1] and outer[0] + outer[2] >= inner[0] + inner[2]
            and outer[1] + outer[3] >= inner[1] + inner[3])


def _solid_fill(sh, theme: Theme) -> Optional[tuple]:
    """(rgb, alpha) of a shape's own solid fill, or None."""
    sppr = sh._element.find(P + "spPr")
    if sppr is None:
        return None
    fill = sppr.find(A + "solidFill")
    if fill is None or not len(fill):
        return None
    return theme.resolve(fill[0])


def text_runs(slide, theme: Theme) -> List[dict]:
    """Every visible text run of a slide whose backdrop is the slide background:
    {box, rgb, need}. Runs on their own fill or over an opaque shape are skipped,
    since a motif behind them cannot be seen."""
    shapes = _flat(slide.shapes)
    out = []
    for i, sh in enumerate(shapes):
        if not _text(sh).strip():
            continue
        b = _box(sh)
        own = _solid_fill(sh, theme)
        if own and own[1] >= 0.99:
            continue
        if any((f := _solid_fill(o, theme)) and f[1] >= 0.99 and _contains(_box(o), b) for o in shapes[:i]):
            continue
        for r in sh._element.iter(A + "r"):
            if not (r.findtext(A + "t") or "").strip():
                continue
            rpr = r.find(A + "rPr")
            fill = rpr.find(A + "solidFill") if rpr is not None else None
            got = theme.resolve(fill[0]) if fill is not None and len(fill) else None
            rgb = got[0] if got else theme.scheme("tx1")
            pt = int(rpr.get("sz", "1800")) / 100 if rpr is not None else 18.0
            bold = rpr is not None and rpr.get("b") == "1"
            large = pt >= 18 or (bold and pt >= 14)
            out.append({"box": b, "rgb": rgb, "need": LARGE_MIN if large else TEXT_MIN, "pt": pt})
    return out


def busy_boxes(slide, W: int, H: int) -> List[Tuple[int, int, int, int]]:
    """Boxes of what the motif should leave room for: text, pictures, logo slots, rules
    and ornaments. Shapes that cover most of the slide (a background band) do not count."""
    out = []
    for sh in _flat(slide.shapes):
        if sh._element.tag == P + "sp" and not _text(sh).strip() and not sh.name.lower().startswith("logo") \
                and _solid_fill_xml(sh) is None and not _has_line(sh):
            continue  # an empty text box draws nothing
        x, y, w, h = _box(sh)
        if w * h > 0.5 * W * H:
            continue
        out.append((x, y, max(w, 1), max(h, 1)))
    return out


def _solid_fill_xml(sh):
    sppr = sh._element.find(P + "spPr")
    return None if sppr is None else sppr.find(A + "solidFill")


def _has_line(sh) -> bool:
    sppr = sh._element.find(P + "spPr")
    ln = None if sppr is None else sppr.find(A + "ln")
    return ln is not None and ln.find(A + "noFill") is None and len(ln)


def deck_reason(slide, conf: dict) -> Optional[str]:
    """None when a deck slide takes the motif, else why not."""
    shapes = _flat(slide.shapes)
    for sh in shapes:
        if getattr(sh, "has_chart", False) or getattr(sh, "has_table", False):
            return "slide with a chart or a table"
        if sh.shape_type == 13 and not sh.name.lower().startswith("logo"):
            return "slide with a picture"
    texts = [sh for sh in shapes if _text(sh).strip()
             and not sh.is_placeholder]
    chars = sum(len(_text(sh)) for sh in texts)
    if len(texts) > conf["max_fields"]:
        return f"{len(texts)} text fields (more than {conf['max_fields']})"
    if chars > conf["max_chars"]:
        return f"{chars} characters of text (more than {conf['max_chars']})"
    sizes = [int(r.get("sz")) / 100 for sh in texts for r in sh._element.iter(A + "rPr") if r.get("sz")]
    if not sizes or max(sizes) < conf["min_pt"]:
        return f"no text set at {conf['min_pt']} pt or more"
    return None


def library_reason(row: dict, conf: dict) -> Optional[str]:
    cat, dens = row.get("category"), row.get("density")
    if cat in NEVER_CATEGORIES or cat not in conf["categories"]:
        return f"category '{cat}' out of scope"
    if dens not in conf["layout_density"]:
        return f"density '{dens}' out of scope"
    return None


# ------------------------------------------------------------------ motifs
# A motif is a list of parts; each part has a bounding box (for the contrast gate)
# ("box", at full strength) and what to draw: a blob ("draw" box, seed) or a segment ("pts").
def _corner_scores(W, H, busy, r):
    corners = {"top-right": (W, 0), "bottom-right": (W, H), "bottom-left": (0, H), "top-left": (0, 0)}
    scores = {}
    for name, (cx, cy) in corners.items():
        zone = (int(cx - r), int(cy - r), int(2 * r), int(2 * r))
        scores[name] = sum(_overlap(zone, b) for b in busy)
    return corners, scores


def _pick_corner(W, H, busy, r, rnd):
    corners, scores = _corner_scores(W, H, busy, r)
    best = min(scores.values())
    ties = sorted(n for n, s in scores.items() if s <= best * 1.05 + 1)
    name = rnd.choice(ties)
    return name, corners[name]


def mist_parts(W, H, busy, dens, rnd) -> List[dict]:
    """A few soft blobs, gathered off the emptiest corner, overlapping like fog."""
    r0 = 0.34 * H
    name, (cx, cy) = _pick_corner(W, H, busy, r0, rnd)
    sx = -1 if cx > 0 else 1  # direction into the slide
    sy = -1 if cy > 0 else 1
    parts = []
    for i in range(dens["blobs"]):
        r = r0 * (1.0 - 0.16 * i) * (0.9 + 0.2 * rnd.random())
        ax = 1.3 + 0.3 * rnd.random()  # wider than tall: fog lies flat
        # spread along the horizontal edge, each a little further in, centres near the edge
        px = cx + sx * (0.15 + 0.6 * i + 0.15 * rnd.random()) * r0
        py = cy + sy * (0.10 + 0.12 * i + 0.10 * rnd.random()) * r0
        w, h = 2 * r * ax, 2 * r
        soft = int(SOFT * min(w, h))
        parts.append({"kind": "blob", "draw": (int(px - w / 2), int(py - h / 2), int(w), int(h)),
                      # the soft edge fades out: the gate looks at the solid core only
                      "box": (int(px - w / 2 + soft), int(py - h / 2 + soft), int(w - 2 * soft), int(h - 2 * soft)),
                      "seed": rnd.randrange(1, 10 ** 6), "alpha": dens["alpha"]})
    return parts


def _seg_box(x0, y0, x1, y1, pad=0):
    # at least 1 EMU thick, so a vertical or horizontal line still overlaps what it crosses
    return (int(min(x0, x1) - pad), int(min(y0, y1) - pad), max(1, int(abs(x1 - x0) + 2 * pad)),
            max(1, int(abs(y1 - y0) + 2 * pad)))


def _emptier_side(W, H, busy) -> bool:
    """True when the right half of the slide holds less than the left half."""
    load_l = sum(_overlap((0, 0, W // 2, H), b) for b in busy)
    load_r = sum(_overlap((W // 2, 0, W - W // 2, H), b) for b in busy)
    return load_r <= load_l


def rain_parts(W, H, busy, dens, rnd) -> List[dict]:
    """Thin streaks at one slant, falling on the emptier side of the slide and thinning
    out towards the content, kept clear of text, rules and pictures."""
    lean = math.radians(12)
    pad = int(0.15 * EMU)
    right = _emptier_side(W, H, busy)
    parts = []
    tries = 0
    while len(parts) < dens["streaks"] and tries < dens["streaks"] * 8:
        tries += 1
        length = (0.35 + 0.45 * rnd.random()) * EMU
        depth = 0.62 * W * rnd.random() ** 1.7  # denser at the edge
        x = W - depth if right else depth
        y = rnd.uniform(-0.1 * H, 1.0 * H)
        x1, y1 = x + length * math.sin(lean), y + length * math.cos(lean)
        bb = _seg_box(x, y, x1, y1, pad)
        if any(_overlap(bb, b) for b in busy) or any(_overlap(bb, q["box"]) for q in parts):
            continue  # clear of the content, and never two streaks side by side
        parts.append({"kind": "seg", "pts": (x, y, x1, y1), "box": _seg_box(x, y, x1, y1),
                      "alpha": dens["alpha"]})
    return parts


def grid_parts(W, H, busy, dens, rnd, marks="lines") -> List[dict]:
    """A fine grid over the emptier half of the slide (left or right), aligned to the
    slide edge, so it reads as one block, not wallpaper."""
    pitch = int(dens["pitch"] * EMU)
    right = _emptier_side(W, H, busy)
    cols = int(0.42 * W // pitch)
    rows = int(H // pitch) + 1
    x0 = W - cols * pitch if right else 0
    parts = []

    # fade towards the content: full strength at the slide edge, a whisper at the
    # inner edge, in a few steps (one freeform per step)
    def fade(i):  # i = columns from the inner edge (0) to the slide edge (cols)
        step = min(4, int(5 * i / max(1, cols)))
        return round(dens["alpha"] * (0.2 + 0.2 * step), 4)

    if marks == "cross":
        arm = int(pitch * 0.12)
        for i in range(cols + 1):
            for j in range(rows + 1):
                x, y = x0 + i * pitch, j * pitch
                a = min(1.0, 2 * fade(i if right else cols - i))
                parts.append({"kind": "seg", "pts": (x - arm, y, x + arm, y), "box": _seg_box(x - arm, y, x + arm, y),
                              "alpha": a})
                parts.append({"kind": "seg", "pts": (x, y - arm, x, y + arm), "box": _seg_box(x, y - arm, x, y + arm),
                              "alpha": a})
        return parts
    # one part per cell side, so the contrast gate drops only the cells under a failing run
    for i in range(cols + 1):
        x = x0 + i * pitch
        k = i if right else cols - i
        for j in range(rows):
            y = j * pitch
            parts.append({"kind": "seg", "pts": (x, y, x, y + pitch), "box": _seg_box(x, y, x, y + pitch),
                          "alpha": fade(k)})
    for j in range(rows + 1):
        y = j * pitch
        for i in range(cols):
            xa = x0 + i * pitch
            k = i + 1 if right else cols - i
            parts.append({"kind": "seg", "pts": (xa, y, xa + pitch, y), "box": _seg_box(xa, y, xa + pitch, y),
                          "alpha": fade(k)})
    return parts


BUILDERS = {"mist": mist_parts, "rain": rain_parts, "grid": grid_parts}


# ------------------------------------------------------------------ contrast gate
WORSE = 0.03  # a run already below its threshold may lose at most this share of its contrast


def gate(parts: List[dict], runs: List[dict], motif_rgb, bg) -> Tuple[List[dict], int]:
    """Drop every part that would take a text run under it below its contrast
    threshold (or, for a run already below it without any motif, lower its contrast
    by more than ``WORSE``). Blobs that overlap each other add up (alpha compositing);
    lines and streaks only cross at points and are judged one by one. Returns
    (kept, dropped)."""
    kept, dropped = [], 0
    base = [contrast(run["rgb"], bg) for run in runs]
    for part in parts:
        ok = True
        for i, run in enumerate(runs):
            if not _overlap(part["box"], run["box"]):
                continue
            show = 1 - part["alpha"]  # share of the background still showing
            if part["kind"] == "blob":
                for k in kept:
                    if k["kind"] == "blob" and _overlap(k["box"], part["box"]) and _overlap(k["box"], run["box"]):
                        show *= 1 - k["alpha"]
            got = contrast(run["rgb"], blend(motif_rgb, 1 - show, bg))
            need = run["need"] if base[i] >= run["need"] else base[i] * (1 - WORSE)
            if got < need:
                ok = False
                break
        if ok:
            kept.append(part)
        else:
            dropped += 1
    return kept, dropped


# ------------------------------------------------------------------ drawing
def _colour(slot: str, alpha: float) -> str:
    return f'<a:schemeClr val="{slot}"><a:alpha val="{int(round(alpha * 100000))}"/></a:schemeClr>'


def _nv(tag: str, sid: int, name: str) -> str:
    inner = "<p:cNvGrpSpPr/>" if tag == "grp" else "<p:cNvSpPr/>"
    return (f'<p:nv{"GrpSpPr" if tag == "grp" else "SpPr"}><p:cNvPr id="{sid}" name="{name}">'
            f'<a:extLst><a:ext uri="{DECORATIVE}"><adec:decorative xmlns:adec="{ADEC}" val="1"/></a:ext>'
            f'</a:extLst></p:cNvPr>{inner}<p:nvPr/></p:nv{"GrpSpPr" if tag == "grp" else "SpPr"}>')


def _xfrm(b, child=False) -> str:
    x, y, w, h = (int(round(v)) for v in b)
    ch = f'<a:chOff x="{x}" y="{y}"/><a:chExt cx="{w}" cy="{h}"/>' if child else ""
    return f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/>{ch}</a:xfrm>'


def _sp(sid, name, b, geom_xml, fill_xml, ln_xml) -> str:
    return (f'<p:sp xmlns:p="{P_NS}" xmlns:a="{A_NS}">{_nv("sp", sid, name)}<p:spPr>{_xfrm(b)}{geom_xml}'
            f'{fill_xml}{ln_xml}</p:spPr></p:sp>')


def _geom_xml(geom, w, h) -> str:
    return etree.tostring(S.custgeom(geom, max(1, int(w)), max(1, int(h)))).decode()


def _lines_geom(segs, b) -> list:
    x0, y0 = b[0], b[1]
    subs = [[("M", a - x0, c - y0), ("L", d - x0, e - y0)] for a, c, d, e in segs]
    return [S.Path(subs, fill=False)]


def build_group(parts: List[dict], motif: str, slot: str, dens: dict, first_id: int) -> etree._Element:
    sid = first_id
    children = []
    boxes = []
    if motif == "mist":
        for n, p in enumerate(parts, 1):
            x, y, w, h = p["draw"]
            geom = S.outline("blob", w, h, strict=False, seed=p["seed"], wobble=0.12)
            sid += 1
            children.append(_sp(sid, f"{motif} {n}", p["draw"], _geom_xml(geom, w, h),
                                f"<a:solidFill>{_colour(slot, p['alpha'])}</a:solidFill>",
                                f'<a:ln><a:noFill/></a:ln><a:effectLst><a:softEdge rad="{int(SOFT * min(w, h))}"/>'
                                '</a:effectLst>'))
            boxes.append(p["draw"])
    else:
        # one freeform per opacity: every streak or grid line of the slide in one shape
        by_alpha: Dict[float, list] = {}
        for p in parts:
            by_alpha.setdefault(p["alpha"], []).append(p["pts"])
        for n, (alpha, segs) in enumerate(sorted(by_alpha.items()), 1):
            xs = [v for s in segs for v in (s[0], s[2])]
            ys = [v for s in segs for v in (s[1], s[3])]
            b = (int(min(xs)), int(min(ys)), max(1, int(max(xs) - min(xs))), max(1, int(max(ys) - min(ys))))
            sid += 1
            ln = (f'<a:ln w="{int(dens["pt"] * S.EMU_PER_PT)}" cap="rnd"><a:solidFill>{_colour(slot, alpha)}'
                  f'</a:solidFill></a:ln>')
            children.append(_sp(sid, f"{motif} {n}", b, _geom_xml(_lines_geom(segs, b), b[2], b[3]),
                                "<a:noFill/>", ln))
            boxes.append(b)
    gx = min(b[0] for b in boxes)
    gy = min(b[1] for b in boxes)
    gw = max(b[0] + b[2] for b in boxes) - gx
    gh = max(b[1] + b[3] for b in boxes) - gy
    xml = (f'<p:grpSp xmlns:p="{P_NS}" xmlns:a="{A_NS}">{_nv("grp", first_id, f"{GROUP_NAME} > {motif}")}'
           f'<p:grpSpPr>{_xfrm((gx, gy, gw, gh), child=True)}</p:grpSpPr>{"".join(children)}</p:grpSp>')
    return etree.fromstring(xml)


def _max_id(slide) -> int:
    return max([int(e.get("id", "0")) for e in slide._element.iter(P + "cNvPr")] + [1])


def apply_slide(slide, conf: dict, seed: int, slide_wh) -> dict:
    """Put the style's motif behind one slide's content. Returns what happened."""
    motif = conf["motif"]
    W, H = (int(v) for v in slide_wh)
    theme = Theme(slide)
    bg = theme.background()
    dark = luminance(bg) < 0.18
    slot = conf["on_dark"] if dark else conf["colour"]
    dens = DENSITY[conf["density"]][motif]
    rnd = random.Random(seed)
    busy = busy_boxes(slide, W, H)
    build = BUILDERS[motif]
    parts = build(W, H, busy, dens, rnd, conf["marks"]) if motif == "grid" else build(W, H, busy, dens, rnd)
    runs = text_runs(slide, theme)
    kept, dropped = gate(parts, runs, theme.scheme(slot), bg)
    item = {"motif": motif, "colour": slot, "background": "dark" if dark else "light", "seed": seed,
            "parts": len(kept), "dropped_by_contrast": dropped}
    if not kept:
        item.update(status="skipped", reason="every part of the motif failed the contrast check" if parts
                    else "no free room for the motif")
        return item
    group = build_group(kept, motif, slot, dens, _max_id(slide) + 1)
    tree = slide.shapes._spTree
    tree.insert(2, group)  # after nvGrpSpPr and grpSpPr: behind every other shape
    item["status"] = "applied"
    return item


# ------------------------------------------------------------------ typography
def set_fonts(prs, typography: Optional[dict]) -> Optional[dict]:
    """Set the theme heading (major) and body (minor) fonts of every master."""
    if not typography:
        return None
    done = set()
    for master in prs.slide_masters:
        part = master.part.part_related_by(RT_THEME)
        if id(part) in done:
            continue
        done.add(id(part))
        root = etree.fromstring(part.blob)
        for key, tag in (("major", "majorFont"), ("minor", "minorFont")):
            lat = root.find(f".//{A}fontScheme/{A}{tag}/{A}latin")
            if key in typography and lat is not None:
                lat.set("typeface", typography[key])
        blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        if hasattr(part, "_element"):
            part._element = etree.fromstring(blob)
        else:
            part._blob = blob
    return {"applied": True, **typography}


# ------------------------------------------------------------------ entry points
def apply_library(prs, lib: str, by_slide: Dict[int, dict], style: dict) -> Optional[dict]:
    """Ambience on the in-scope layouts of one library (visuals never). None when the
    style has no ambience."""
    conf = style.get("ambience")
    if not conf:
        return None
    items = []
    if conf["motif"]:
        for n, slide in enumerate(prs.slides, 1):
            row = by_slide.get(n)
            if row is None:
                continue
            reason = "visuals never take an ambience" if lib == "visuals" else library_reason(row, conf)
            if reason:
                items.append({"id": row["id"], "status": "out-of-scope", "reason": reason})
                continue
            seed = slide_seed(conf, style.get("seed", 0), row["id"])
            items.append({"id": row["id"], **apply_slide(slide, conf, seed, (prs.slide_width, prs.slide_height))})
    return {"motif": conf["motif"], "density": conf["density"], "typography": set_fonts(prs, conf["typography"]),
            "items": items}


def apply_deck(prs, style: dict) -> Optional[dict]:
    conf = style.get("ambience")
    if not conf:
        return None
    items = []
    if conf["motif"]:
        for n, slide in enumerate(prs.slides, 1):
            key = f"slide {n}"
            reason = deck_reason(slide, conf)
            if reason:
                items.append({"id": key, "status": "out-of-scope", "reason": reason})
                continue
            seed = slide_seed(conf, style.get("seed", 0), key)
            items.append({"id": key, **apply_slide(slide, conf, seed, (prs.slide_width, prs.slide_height))})
    return {"motif": conf["motif"], "density": conf["density"], "typography": set_fonts(prs, conf["typography"]),
            "items": items}
