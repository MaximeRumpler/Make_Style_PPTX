"""The skill folder must stay a valid, self-contained Agent Skill."""
import json
import re
from pathlib import Path

import pytest

import restyle as R

ROOT = Path(__file__).resolve().parents[1]
SKILL = ROOT / "skills" / "make-style-pptx"


def _frontmatter(text):
    m = re.match(r"^---\n(.*?)\n---\n", text, re.S)
    assert m, "SKILL.md must start with YAML frontmatter"
    out = {}
    for line in m.group(1).splitlines():
        if re.match(r"^[A-Za-z_-]+:", line):
            k, _, v = line.partition(":")
            out[k.strip()] = v.strip()
    return out


def test_skill_md():
    text = (SKILL / "SKILL.md").read_text(encoding="utf-8")
    fm = _frontmatter(text)
    assert fm["name"] == SKILL.name == "make-style-pptx"
    assert 50 < len(fm["description"]) <= 1024
    assert "change le style" in fm["description"]  # French triggers
    assert text.count("\n") < 160


def test_referenced_files_exist():
    text = (SKILL / "SKILL.md").read_text(encoding="utf-8")
    for rel in set(re.findall(r"`((?:references|scripts|assets|styles)/[\w./-]+)`", text)):
        assert (SKILL / rel).exists(), f"SKILL.md mentions {rel} but it does not exist"


def test_styles_in_skill_md_exist():
    """The style table lists exactly the styles that ship."""
    text = (SKILL / "SKILL.md").read_text(encoding="utf-8")
    listed = set(re.findall(r"^\| `(\w+)` \|", text, re.M)) - {"stagger", "cascade", "arc", "deconstructed"}
    assert listed == {p.stem for p in (SKILL / "styles").glob("*.json")}


def test_flags_in_skill_md_exist():
    """Every flag SKILL.md uses with a script is one the script accepts."""
    text = (SKILL / "SKILL.md").read_text(encoding="utf-8")
    for script in ("restyle.py", "restyle_sheet.py"):
        src = (SKILL / "scripts" / script).read_text(encoding="utf-8")
        accepted = set(re.findall(r'add_argument\("(--[\w-]+)"', src))
        for line in re.findall(rf"scripts/{script} (.*)", text):
            for flag in re.findall(r"(?<![\w-])(--[\w-]+)", line):
                assert flag in accepted, (script, flag)


STDLIB = {"__future__", "argparse", "collections", "colorsys", "copy", "csv", "functools", "hashlib", "io", "itertools", "json",
          "math", "os", "pathlib", "random", "re", "shutil", "struct", "subprocess", "sys", "tempfile", "typing",
          "xml", "zipfile", "zlib"}
THIRD_PARTY = {"pptx", "PIL", "lxml"}  # python-pptx, Pillow, lxml: nothing else to install


def test_self_contained_and_light():
    for py in (SKILL / "scripts").glob("*.py"):
        src = py.read_text(encoding="utf-8")
        assert "parent.parent.parent" not in src, py.name
        for a, b in re.findall(r"^\s*(?:from (\w+)[\w.]* import |import (\w+))", src, re.M):
            mod = a or b
            local = (SKILL / "scripts" / (mod + ".py")).is_file()
            assert mod in STDLIB | THIRD_PARTY or local, (py.name, mod)
    reqs = (SKILL / "requirements.txt").read_text().split()
    assert {r.split(">=")[0] for r in reqs} == {"python-pptx", "Pillow", "lxml"}
    for name in ("LICENSE.txt", "NOTICE.txt"):
        assert (SKILL / name).is_file(), name


def test_plugin_manifests_agree():
    plugin = json.loads((ROOT / ".claude-plugin" / "plugin.json").read_text(encoding="utf-8"))
    market = json.loads((ROOT / ".claude-plugin" / "marketplace.json").read_text(encoding="utf-8"))
    fm = _frontmatter((SKILL / "SKILL.md").read_text(encoding="utf-8"))
    assert plugin["name"] == market["plugins"][0]["name"] == fm["name"]
    version = re.search(r'version: "([\d.]+)"', (SKILL / "SKILL.md").read_text(encoding="utf-8")).group(1)
    assert plugin["version"] == market["plugins"][0]["version"] == version


@pytest.mark.parametrize("style", ["organic", "paper"])
def test_sample_deck_restyles(style, tmp_path):
    """The bundled sample deck is a real Make_Clean deck: each style changes it."""
    out, report = R.run(SKILL / "assets" / "sample-deck.pptx", R.load_style(style), tmp_path / "s.pptx", False)
    assert out.is_file()
    assert report["summary"]["total"].get("restyled", 0) >= 3
