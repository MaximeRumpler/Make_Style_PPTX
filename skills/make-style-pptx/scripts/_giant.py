"""Giant typography: one huge word or figure that carries a slide or sets its tone.

The agent names the words, slide by slide, from what each slide says (``--giant 3=2027,7=TRUST``
or ``giant.words`` in a style file); nothing is invented here. A style opts in with a ``giant`` key::

    "giant": {"mode": "background", "colour": "tx2", "words": {"3": "2027"}}

Two modes, deck only:

- ``background``: the word is set in the theme heading font, very large, as a faint tint of a
  theme colour behind the content, bleeding off the left and bottom edges by fixed offsets. It is
  one text box ``giant#1 > background``, first in the shape tree, marked decorative (no alt text,
  skipped by screen readers). Text contrast is measured with the word behind it (WCAG 2.x, the
  same gate as the ambience); the word is lightened in steps, then dropped.
- ``foreground``: the slide's own shape whose text is exactly the word is set at a fixed share of
  the slide width, so the word is the message. It stays where it is, read first as before. When
  the larger box would touch other text, leave the slide or run off it, the slide is left alone.

Rules: a word has 8 characters at most and must appear in the slide's own text (a word that is
not on the slide is decoration, refused); at most one per slide; about 3 slides in 10 at most;
slides with a chart, a table or a picture are left alone. Sizes come from conservative width
factors for a bold display face, no auto-shrink below ``min_pt``: under it the slide is left
alone and keeps its normal title.
"""
from __future__ import annotations

import math
from typing import Dict, List, Optional

from lxml import etree

import _ambience as AMB

A, P, A_NS, P_NS = AMB.A, AMB.P, AMB.A_NS, AMB.P_NS
EMU_PT = 12700
MAX_CHARS = 8
CHAR_EM = 0.68  # width of one capital in a bold display face, in em (conservative)
LINE = 1.2  # line height in em
GROUP_NAME = "giant#1"
DEFAULTS = {"mode": "background", "colour": "tx2", "on_dark": "bg1", "share": 0.7, "min_pt": 72,
            "bleed": {"left": 0.03, "bottom": 0.10}, "alphas": [0.10, 0.07, 0.05], "per10": 3, "words": {}}


class GiantError(ValueError):
    pass


def normalise(conf) -> Optional[dict]:
    """The ``giant`` value of a style, validated, with defaults. None = no giant typography."""
    if not conf:
        return None
    if not isinstance(conf, dict):
        raise GiantError("giant must be an object")
    unknown = set(conf) - set(DEFAULTS)
    if unknown:
        raise GiantError(f"giant: unknown key(s) {', '.join(sorted(unknown))}")
    out = {**DEFAULTS, **conf}
    if out["mode"] not in ("background", "foreground"):
        raise GiantError("giant: mode must be 'background' or 'foreground'")
    for key in ("colour", "on_dark"):
        if out[key] not in AMB.SCHEME:
            raise GiantError(f"giant: {key} must be a theme colour ({', '.join(AMB.SCHEME)})")
    if not (isinstance(out["share"], (int, float)) and 0.2 <= out["share"] <= 1.0):
        raise GiantError("giant: share is the part of the slide width, between 0.2 and 1")
    if not (isinstance(out["min_pt"], (int, float)) and out["min_pt"] >= 24):
        raise GiantError("giant: min_pt must be 24 or more")
    if not (isinstance(out["per10"], (int, float)) and 0 < out["per10"] <= 10):
        raise GiantError("giant: per10 must be between 0 and 10")
    words = out["words"]
    if not isinstance(words, dict):
        raise GiantError("giant: words is {\"<slide number>\": \"<word>\"}")
    clean = {}
    for key, word in words.items():
        if not str(key).isdigit() or not isinstance(word, str) or not word.strip():
            raise GiantError("giant: words is {\"<slide number>\": \"<word>\"}")
        clean[int(key)] = word.strip()
    out["words"] = clean
    return out


def parse_words(raw: str) -> Dict[int, str]:
    """``3=2027,7=TRUST`` -> {3: "2027", 7: "TRUST"}."""
    out = {}
    for part in (raw or "").split(","):
        if not part.strip():
            continue
        num, sep, word = part.partition("=")
        if not sep or not num.strip().isdigit() or not word.strip():
            raise GiantError(f"giant words are slide=word pairs, got '{part}'")
        out[int(num)] = word.strip()
    return out


def _slide_text(slide) -> str:
    return "\n".join(AMB._text(sh) for sh in AMB._flat(slide.shapes))


def _reason(slide, word: str) -> Optional[str]:
    if len(word) > MAX_CHARS:
        return f"'{word}' is longer than {MAX_CHARS} characters; the slide keeps its normal title"
    if word.lower() not in _slide_text(slide).lower():
        return f"'{word}' is not in the slide's own text (a giant word comes from the message)"
    for sh in AMB._flat(slide.shapes):
        if getattr(sh, "has_chart", False) or getattr(sh, "has_table", False):
            return "slide with a chart or a table"
        if sh.shape_type == 13 and not sh.name.lower().startswith("logo"):
            return "slide with a picture"
    return None


def _size_pt(word: str, width_pt: float, height_pt: float, share: float) -> float:
    """Font size so the word fills ``share`` of the width, never taller than 0.8 of the slide, in whole points."""
    return float(math.floor(min(share * width_pt / (len(word) * CHAR_EM), 0.8 * height_pt / LINE)))  # whole points


def _scheme(slot: str, alpha: Optional[float] = None) -> str:
    inner = f'<a:alpha val="{int(round(alpha * 100000))}"/>' if alpha is not None else ""
    return f'<a:schemeClr val="{slot}">{inner}</a:schemeClr>'


def _background_xml(word: str, box, size_pt: float, slot: str, alpha: float, sid: int) -> etree._Element:
    x, y, w, h = (int(v) for v in box)
    xml = (f'<p:sp xmlns:p="{P_NS}" xmlns:a="{A_NS}">{AMB._nv("sp", sid, f"{GROUP_NAME} > background")}'
           f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
           '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
           '<p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit/></a:bodyPr>'
           '<a:lstStyle/><a:p><a:pPr algn="l"/>'
           f'<a:r><a:rPr lang="en-US" sz="{int(size_pt * 100)}" b="1" dirty="0"><a:solidFill>{_scheme(slot, alpha)}'
           f'</a:solidFill><a:latin typeface="+mj-lt"/></a:rPr><a:t>{_esc(word)}</a:t>'
           '</a:r></a:p></p:txBody></p:sp>')
    return etree.fromstring(xml)


def _esc(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _background(slide, word: str, conf: dict, slide_wh) -> dict:
    W, H = (int(v) for v in slide_wh)
    theme = AMB.Theme(slide)
    bg = theme.background()
    slot = conf["on_dark"] if AMB.luminance(bg) < 0.18 else conf["colour"]
    size = _size_pt(word, W / EMU_PT, H / EMU_PT, conf["share"])
    if size < conf["min_pt"]:
        return {"status": "kept", "reason": f"would need {size:.0f} pt, under the style's {conf['min_pt']} pt minimum"}
    bleed = conf["bleed"]
    bw, bh = int(len(word) * CHAR_EM * size * EMU_PT), int(size * LINE * EMU_PT)
    x = int(-bleed["left"] * W)
    y = int(H - bh + bleed["bottom"] * H)  # the foot of the letters runs off the bottom edge
    box = (x, y, bw, bh)
    runs = AMB.text_runs(slide, theme)
    inside = (max(x, 0), max(y, 0), min(x + bw, W) - max(x, 0), min(y + bh, H) - max(y, 0))
    for alpha in conf["alphas"]:
        part = {"kind": "blob", "box": inside, "alpha": alpha}
        kept, dropped = AMB.gate([part], runs, theme.scheme(slot), bg)
        if kept:
            sp = _background_xml(word, box, size, slot, alpha, AMB._max_id(slide) + 1)
            slide.shapes._spTree.insert(2, sp)  # behind every other shape
            return {"status": "applied", "mode": "background", "word": word, "size_pt": round(size),
                    "colour": slot, "alpha": alpha, "lightened": alpha != conf["alphas"][0]}
    return {"status": "skipped", "reason": "text contrast over the word stays below 4.5:1 at the lightest tint"}


def _foreground(slide, word: str, conf: dict, slide_wh) -> dict:
    W, H = (int(v) for v in slide_wh)
    target = [sh for sh in AMB._flat(slide.shapes) if AMB._text(sh).strip().lower() == word.lower()]
    if len(target) != 1:
        return {"status": "kept", "reason": f"no single shape holds exactly '{word}' (found {len(target)})"}
    sh = target[0]
    size = _size_pt(word, W / EMU_PT, H / EMU_PT, conf["share"])
    if size < conf["min_pt"]:
        return {"status": "kept", "reason": f"would need {size:.0f} pt, under the style's {conf['min_pt']} pt minimum"}
    x, y = int(sh.left or 0), int(sh.top or 0)
    new = (x, y, int(len(word) * CHAR_EM * size * EMU_PT), int(size * LINE * EMU_PT))
    if new[0] + new[2] > W or new[1] + new[3] > H:
        return {"status": "kept", "reason": "the larger word would run off the slide"}
    for other in AMB._flat(slide.shapes):
        if other._element is sh._element or not AMB._text(other).strip():
            continue
        if AMB._overlap(new, AMB._box(other)):
            return {"status": "kept", "reason": f"the larger word would touch '{other.name}'"}
    for rpr in sh._element.iter(A + "rPr"):
        rpr.set("sz", str(int(size * 100)))
    for rpr in sh._element.iter(A + "endParaRPr"):
        rpr.set("sz", str(int(size * 100)))
    sh.left, sh.top, sh.width, sh.height = new
    body = sh._element.find(P + "txBody/" + A + "bodyPr")
    if body is not None:
        body.set("wrap", "none")
        for child in list(body):
            if etree.QName(child).localname in ("normAutofit", "spAutoFit", "noAutofit"):
                body.remove(child)
        etree.SubElement(body, A + "noAutofit")
    return {"status": "applied", "mode": "foreground", "word": word, "size_pt": round(size), "shape": sh.name}


def apply_deck(prs, style: dict, skip: Optional[set] = None) -> Optional[dict]:
    """Giant words on the named slides of a deck, in slide order, within the per-10 share.
    ``skip`` holds slide numbers that already carry an expressive device."""
    conf = style.get("giant")
    if not conf or not conf["words"]:
        return None
    n = len(prs.slides)
    budget = max(1, int(n * conf["per10"] // 10))
    skip = skip or set()
    items: List[dict] = []
    used = 0
    for num in sorted(conf["words"]):
        item = {"id": f"slide {num}", "word": conf["words"][num]}
        items.append(item)
        if not 1 <= num <= n:
            item.update(status="kept", reason=f"the deck has {n} slides")
            continue
        slide = prs.slides[num - 1]
        why = _reason(slide, item["word"])
        if why:
            item.update(status="kept", reason=why)
        elif num in skip:
            item.update(status="kept", reason="the slide already carries an expressive device (one per slide)")
        elif used >= budget:
            item.update(status="kept", reason=f"dosage: {budget} giant slide{'s' if budget > 1 else ''} at most "
                                              f"for {n} slides")
        else:
            run = _foreground if conf["mode"] == "foreground" else _background
            item.update(run(slide, item["word"], conf, (prs.slide_width, prs.slide_height)))
            used += item["status"] == "applied"
    return {"mode": conf["mode"], "budget": budget, "items": items}
