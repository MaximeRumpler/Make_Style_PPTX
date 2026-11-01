"""Style from an example deck: patterns only, a valid draft, nothing copied."""
import json
import shutil
from pathlib import Path

from pptx import Presentation

import restyle as R
import style_from_deck as SD

ROOT = Path(__file__).resolve().parent.parent
SAMPLE = ROOT / "skills" / "make-style-pptx" / "assets" / "sample-deck.pptx"


def restyled(tmp_path, style):
    out, _ = R.run(SAMPLE, R.load_style(style), tmp_path / f"{style}.pptx")
    return out


def test_plain_deck_stays_sober():
    a = SD.analyse(Presentation(str(SAMPLE)))
    assert a["kinds"]["rectangle"] > 0.9 and a["rotated"] == 0
    d = SD.draft(a)
    assert d["style"]["roles"] == {} and d["style"]["ambience"] is None
    assert any("sober" in line for line in d["rationale"])


def test_organic_and_paper_decks_are_recognised(tmp_path):
    org = SD.draft(SD.analyse(Presentation(str(restyled(tmp_path, "organic")))))
    kinds = {c["kind"] for r in org["style"]["roles"].values() for c in r["shapes"]}
    assert "blob" in kinds and (org["style"]["ambience"] or {}).get("motif", "wash") in ("wash", "mist")
    pap = SD.draft(SD.analyse(Presentation(str(restyled(tmp_path, "paper")))))
    assert {c["kind"] for r in pap["style"]["roles"].values() for c in r["shapes"]} == {"handdrawn"}


def test_analysis_keeps_no_text_names_or_colours(tmp_path):
    prs = Presentation(str(restyled(tmp_path, "organic")))
    text = " ".join(sh.text_frame.text for s in prs.slides for sh in s.shapes if sh.has_text_frame).split()
    analysis = SD.analyse(prs)
    blob = json.dumps(analysis) + json.dumps(SD.draft(analysis)["style"]["roles"])
    keys = {k for k in SD.analyse(prs)} | {"filled", "shadow", "kinds"}  # our own vocabulary, not the deck's
    assert not [w for w in set(text) if len(w) > 5 and w.lower() not in keys and w.lower() in blob.lower()]
    assert "srgbClr" not in blob and "#" not in blob.replace("#1", "")


def test_cli_writes_a_valid_draft_and_refuses_to_overwrite(tmp_path):
    out = tmp_path / "styles" / "mine.json"
    assert SD.main([str(restyled(tmp_path, "organic")), "--name", "My Look", "--out", str(out)]) == 0
    assert json.loads(out.read_text())["name"] == "my-look"
    assert R.load_style(str(out))["name"] == "my-look"
    assert SD.main([str(SAMPLE), "--out", str(out)]) == 2
    assert SD.main(["--check", str(out)]) == 0
    bad = tmp_path / "bad.json"
    bad.write_text(json.dumps({"name": "bad", "roles": {"nope": {}}}))
    assert SD.main(["--check", str(bad)]) == 2
    assert SD.main([str(tmp_path / "missing.pptx")]) == 2
