"""Rebuild the vendored Make_Clean subset used by the tests.

    python tests/fixtures/make_fixture.py <Make_Clean_PPTX clone at tag v1.1.0>

Keeps a few layouts and visuals of the bundled libraries (low-text slides with
containers, a cover with the logo slot, a table and a chart that must stay
untouched), so the fixture stays small. Same ids, fields and limits as v1.1.0.
A second, smaller set holds the slides the ambience layer works on (dark and light
covers, a divider, a key message, a quote) and a few it must leave alone.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "skills" / "make-style-pptx" / "scripts"))
import restyle_sheet  # noqa: E402

VERSION = "1.1.0"
IDS = ["L-01", "L-05", "L-24", "L-38", "L-41", "L-47", "L-50", "L-54",
       "V-CMP-01", "V-KPI-01", "V-PRO-04", "V-TIM-05"]
OUT = Path(__file__).resolve().parent / f"make-clean-{VERSION}"
AMBIENCE_IDS = ["L-02", "L-09", "L-12", "L-14", "L-24", "L-37", "L-41", "V-TIM-05"]
AMBIENCE_OUT = Path(__file__).resolve().parent / f"make-clean-{VERSION}-ambience"

if __name__ == "__main__":
    src = Path(sys.argv[1]) / "skills" / "make-clean-pptx" / "assets"
    for out, ids in ((OUT, IDS), (AMBIENCE_OUT, AMBIENCE_IDS)):
        kept = restyle_sheet.subset(src, out, set(ids))
        print(f"{out}: {', '.join(kept)}")
