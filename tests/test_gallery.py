"""Reference styles and the gallery: every style loads, is described, and restyles the sample deck cleanly."""
import json
from pathlib import Path

import pytest

import gallery as G
import restyle as R

SKILL = Path(__file__).resolve().parent.parent / "skills" / "make-style-pptx"
NAMES = sorted(p.stem for p in (SKILL / "styles").glob("*.json"))


def test_every_style_is_in_the_catalog_and_families_are_known():
    cat = G.catalog()
    assert set(cat) == set(NAMES) and len(NAMES) >= 15
    assert {c["family"] for c in cat.values()} == set(G.FAMILIES)
    assert all(c["mode"] in ("light", "dark", "either") and c["use_when"] for c in cat.values())


@pytest.mark.parametrize("name", NAMES)
def test_style_restyles_the_sample_deck_without_quality_warnings(name, tmp_path):
    st = R.load_style(name)
    assert st["description"] and "#" not in json.dumps(st["roles"]).replace("#1", "")
    _, result = G.run_style(G.SAMPLE, name, tmp_path)
    assert result["quality"] == []
    assert sum(result["summary"].values()) > 0


def test_styles_use_theme_colours_only():
    for name in NAMES:
        text = (SKILL / "styles" / f"{name}.json").read_text()
        assert "srgb" not in text.lower() and not any(ch in text for ch in ("#FF", "#ff"))


def test_gallery_writes_table_and_report(tmp_path):
    out = tmp_path / "gal"
    assert G.main(["--styles", "clean,organic,rain", "--out", str(out), "--dpi", "20"]) == 0
    md = (out / "gallery.md").read_text()
    assert md.count("\n| `") == 3 and "| Sober |" in md and "ambience: rain" in md
    report = json.loads((out / "gallery.json").read_text())
    assert set(report) == {"clean", "organic", "rain"} and all(r["quality"] == [] for r in report.values())
    assert G.main(["--styles", "nope", "--out", str(tmp_path / "x")]) == 2
