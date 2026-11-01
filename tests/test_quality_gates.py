"""Quality gates: each one fires on a deck built to break it, and a clean styled deck passes."""
import copy
import json
import shutil
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

import _ambience as AMB
import quality_gates as QG
import restyle as R

FIXTURE = Path(__file__).resolve().parent / "fixtures" / "make-clean-1.1.0-ambience"
A = AMB.A


def deck(tmp_path):
    out = tmp_path / "d.pptx"
    shutil.copyfile(FIXTURE / "layouts.pptx", out)
    return Presentation(str(out))


def gates(prs, **kw):
    return [(w["slide"], w["gate"]) for w in QG.check(prs, **kw)["warnings"]]


def shape(prs, n, name):
    return next(sh for sh in prs.slides[n - 1].shapes if sh.name == name)


def test_clean_deck_passes(tmp_path):
    assert QG.check(deck(tmp_path))["warnings"] == []


def test_overlap(tmp_path):
    prs = deck(tmp_path)
    t, s = shape(prs, 5, "Title"), shape(prs, 5, "Kicker")
    s.left, s.top, s.width, s.height = t.left, t.top, t.width, t.height
    assert (5, "overlap") in gates(prs)


def test_stacked_contrast_with_a_heavy_translucent_layer(tmp_path):
    prs = deck(tmp_path)
    slide = prs.slides[2]
    heavy = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    heavy.fill.solid()
    heavy.fill.fore_color.rgb = __import__("pptx.dml.color", fromlist=["RGBColor"]).RGBColor(0x80, 0x80, 0x80)
    clr = heavy._element.find(".//" + A + "srgbClr")
    etree_alpha = clr.makeelement(A + "alpha", {"val": "60000"})
    clr.append(etree_alpha)
    slide.shapes._spTree.remove(heavy._element)
    slide.shapes._spTree.insert(2, heavy._element)  # behind the text
    got = [w for w in QG.check(prs)["warnings"] if w["gate"] == "contrast"]
    assert got and all(w["slide"] == 3 for w in got) and "over the stack" in got[0]["message"]


def test_reading_order_title_first_and_decorative_flag(tmp_path):
    prs = deck(tmp_path)
    tree = prs.slides[4].shapes._spTree
    title = shape(prs, 5, "Title")._element
    tree.remove(title)
    tree.append(title)  # now after the body text
    assert (5, "reading-order") in gates(prs)
    prs = deck(tmp_path)
    slide = prs.slides[2]
    box = slide.shapes.add_textbox(0, 0, 100, 100)
    box.name = "giant#1 > background"
    assert (3, "reading-order") in gates(prs)


def test_effects_budget_and_undeclared_gradient(tmp_path):
    prs = deck(tmp_path)
    slide = prs.slides[3]
    sp = slide.shapes.add_shape(1, 0, 0, 1000, 1000)
    sp.rotation = 7
    sp.fill.gradient()
    from lxml import etree
    sp._element.spPr.append(etree.fromstring(
        "<a:effectLst xmlns:a='%s'><a:outerShdw/><a:glow rad='1'/><a:softEdge rad='1'/></a:effectLst>" % A[1:-1]))
    got = QG.check(prs, budget=3)["warnings"]
    effects = [w["message"] for w in got if w["gate"] == "effects"]
    assert any("kinds of effect" in m for m in effects)
    assert any("gradient is not declared" in m for m in effects) and any("glow is not declared" in m for m in effects)
    style = R.load_style("organic")  # mist declares transparency and soft edge only
    assert any("gradient" in w["message"] for w in QG.check(prs, style)["warnings"])
    leak = {"name": "l", "seed": 1, "roles": {}, "ambience": {"motif": "leak"}}
    p = tmp_path / "l.json"
    p.write_text(json.dumps(leak))
    declared = [w["message"] for w in QG.check(prs, R.load_style(str(p)))["warnings"] if w["gate"] == "effects"]
    assert not any("gradient is not declared" in m for m in declared)


def test_dosage(tmp_path):
    prs = deck(tmp_path)
    for n in (2, 3, 5):
        prs.slides[n - 1]._element.cSld.set("name", "V-ORB-01-4")
    got = gates(prs)
    assert (3, "dosage") in got and (0, "dosage") in got


def test_restyled_deck_report_carries_the_gates_and_the_cli(tmp_path):
    src = tmp_path / "in.pptx"
    shutil.copyfile(FIXTURE / "layouts.pptx", src)
    out, report = R.run(src, R.load_style("organic"), tmp_path / "out.pptx")
    assert report["quality"] == {"slides": 7, "warnings": [], "by_gate": {}}
    assert QG.main([str(out), "--style", "organic", "--strict"]) == 0
    prs = Presentation(str(out))
    t, s = shape(prs, 5, "Title"), shape(prs, 5, "Kicker")
    s.left, s.top = t.left, t.top
    bad = tmp_path / "bad.pptx"
    prs.save(str(bad))
    assert QG.main([str(bad)]) == 0 and QG.main([str(bad), "--strict"]) == 1
    assert QG.main([str(tmp_path / "missing.pptx")]) == 2
