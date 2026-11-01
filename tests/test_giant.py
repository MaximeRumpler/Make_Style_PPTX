"""Giant typography: words come from the slide, background tint gated by contrast, foreground enlarged, dosage."""
import json
import shutil
from pathlib import Path

import pytest
from pptx import Presentation

import _ambience as AMB
import _giant as GI
import restyle as R

FIXTURE = Path(__file__).resolve().parent / "fixtures" / "make-clean-1.1.0-ambience"
A = AMB.A


def plain_style(tmp_path, **giant):
    p = tmp_path / "giant.json"
    p.write_text(json.dumps({"name": "giant", "seed": 3, "roles": {}, "giant": giant or None}))
    return R.load_style(str(p))


def deck(tmp_path, texts=None):
    """The fixture library as a deck; ``texts`` = {slide number: new text of its Number/Title shape}."""
    out = tmp_path / "deck.pptx"
    shutil.copyfile(FIXTURE / "layouts.pptx", out)
    if texts:
        prs = Presentation(str(out))
        for n, (shape, text) in texts.items():
            next(sh for sh in prs.slides[n - 1].shapes if sh.name == shape).text_frame.paragraphs[0].runs[0].text = text
        prs.save(str(out))
    return out


def giant_report(tmp_path, style, **kw):
    out, report = R.run(deck(tmp_path, kw.get("texts")), style, tmp_path / "out.pptx", force=True)
    return out, {it["id"]: it for it in report.get("giant", {}).get("items", [])}, report


def test_background_word_behind_everything(tmp_path):
    st = plain_style(tmp_path, words={"3": "2027"})
    out, items, _ = giant_report(tmp_path, st, texts={3: ("Message", "2027")})
    assert items["slide 3"]["status"] == "applied" and items["slide 3"]["size_pt"] >= 72
    slide = Presentation(str(out)).slides[2]
    first = slide.shapes[0]
    assert first.name == "giant#1 > background" and first.text_frame.text == "2027"
    xml = first._element
    assert xml.find(f".//{{{AMB.ADEC}}}decorative") is not None  # decorative: no alt text, skipped by readers
    assert xml.find(".//" + A + "latin").get("typeface") == "+mj-lt"  # the theme heading font
    clr = xml.find(".//" + A + "schemeClr")
    assert clr.get("val") in AMB.SCHEME and int(clr.find(A + "alpha").get("val")) <= 10000
    assert not list(xml.iter(A + "srgbClr"))
    assert first.left < 0  # bleeds off the left edge by the style's offset


def test_the_word_must_come_from_the_slide_and_be_short(tmp_path):
    st = plain_style(tmp_path, words={"3": "TRUST", "99": "2027"})
    _, items, _ = giant_report(tmp_path, st)
    assert "not in the slide's own text" in items["slide 3"]["reason"]
    assert items["slide 99"]["status"] == "kept"
    assert GI._reason(Presentation(str(deck(tmp_path))).slides[2], "WAYTOOLONG").startswith("'WAYTOOLONG' is longer")


def test_chart_slides_and_dosage(tmp_path):
    ids = {r["id"]: int(r["slide"]) for r in R.read_index(FIXTURE / "layouts.csv")[1]}
    chart = ids["L-41"]
    texts = {3: ("Message", "2027"), chart: ("Title", "2027")}
    words = {"3": "2027", str(chart): "2027"}
    st = plain_style(tmp_path, words=words)
    _, items, report = giant_report(tmp_path, st, texts=texts)
    assert items[f"slide {chart}"]["status"] == "kept"
    assert report["giant"]["budget"] == 2  # 7 slides, 3 in 10
    st = plain_style(tmp_path, words={"3": "2027", "5": "2027", "6": "2027"}, per10=1)
    texts = {n: ("Title", "2027") for n in (5, 6)}
    texts[3] = ("Message", "2027")
    _, items, _ = giant_report(tmp_path, st, texts=texts)
    applied = [k for k, v in items.items() if v["status"] == "applied"]
    assert len(applied) == 1 and any(v.get("reason", "").startswith("dosage") for v in items.values())


def test_contrast_gate_lightens_then_drops(tmp_path, monkeypatch):
    st = plain_style(tmp_path, words={"3": "2027"})
    seen = []
    real = AMB.gate

    def spy(parts, runs, rgb, bg):
        seen.append(parts[0]["alpha"])
        return ([], 1) if len(seen) < 3 else real(parts, runs, rgb, bg)

    monkeypatch.setattr(AMB, "gate", spy)
    _, items, _ = giant_report(tmp_path, st, texts={3: ("Message", "2027")})
    assert seen == [0.10, 0.07, 0.05] and items["slide 3"]["alpha"] == 0.05 and items["slide 3"]["lightened"]
    monkeypatch.setattr(AMB, "gate", lambda *a: ([], 1))
    _, items, _ = giant_report(tmp_path, st, texts={3: ("Message", "2027")})
    assert items["slide 3"]["status"] == "skipped" and "contrast" in items["slide 3"]["reason"]


def test_foreground_enlarges_the_own_shape_or_leaves_it(tmp_path):
    texts = {5: ("Title", "2027")}
    st = plain_style(tmp_path, mode="foreground", words={"5": "2027"}, share=0.25, min_pt=40)
    out, items, _ = giant_report(tmp_path, st, texts=texts)
    assert items["slide 5"]["status"] == "applied" and items["slide 5"]["shape"] == "Title"
    title = next(sh for sh in Presentation(str(out)).slides[4].shapes if sh.name == "Title")
    assert int(title._element.find(".//" + A + "rPr").get("sz")) == items["slide 5"]["size_pt"] * 100
    # at 70 % of the width the word would run into the body below: the slide is left alone
    st = plain_style(tmp_path, mode="foreground", words={"5": "2027"})
    _, items, _ = giant_report(tmp_path, st, texts=texts)
    assert items["slide 5"]["status"] == "kept" and ("would touch" in items["slide 5"]["reason"]
                                                      or "run off" in items["slide 5"]["reason"])


def test_below_the_minimum_size_the_slide_keeps_its_title(tmp_path):
    st = plain_style(tmp_path, words={"3": "2027"}, share=0.2, min_pt=200)
    _, items, _ = giant_report(tmp_path, st, texts={3: ("Message", "2027")})
    assert items["slide 3"]["status"] == "kept" and "minimum" in items["slide 3"]["reason"]


def test_expressive_slide_gets_no_second_device(tmp_path):
    d = deck(tmp_path, {3: ("Message", "2027")})
    prs = Presentation(str(d))
    prs.slides[2]._element.cSld.set("name", "V-SIG-01-1")
    prs.save(str(d))
    out, report = R.run(d, plain_style(tmp_path, words={"3": "2027"}), tmp_path / "x.pptx")
    assert "one per slide" in report["giant"]["items"][0]["reason"]


def test_cli_and_determinism(tmp_path):
    d = deck(tmp_path, {3: ("Message", "2027")})
    args = [str(d), "--style", "organic", "--giant", "3=2027", "--force"]
    assert R.main(args + ["--out", str(tmp_path / "a.pptx")]) == 0
    assert R.main(args + ["--out", str(tmp_path / "b.pptx")]) == 0
    assert (tmp_path / "a.pptx").read_bytes() == (tmp_path / "b.pptx").read_bytes()
    assert Presentation(str(tmp_path / "a.pptx")).slides[2].shapes[0].name == "giant#1 > background"
    assert R.main([str(d), "--style", "organic", "--giant", "two=2027", "--out", str(tmp_path / "c.pptx")]) == 2


def test_defaults_add_nothing_and_bad_config_refused(tmp_path):
    out, report = R.run(deck(tmp_path), plain_style(tmp_path), tmp_path / "n.pptx")
    assert "giant" not in report
    for bad in ({"mode": "huge"}, {"share": 3}, {"words": {"x": "1"}}, {"colour": "pink"}, {"nope": 1}):
        with pytest.raises(R.RestyleError):
            plain_style(tmp_path, **bad)
