"""Arrangement layer: positions per count, grid rules, overlap, margins, reading order, fallback, determinism."""
import json
import math

import pytest
from pptx import Presentation

import _arrange as AR
import restyle as R
from test_restyle import FIXTURE, items, shapes_by_id

E = AR.EMU


def style_with(name, base="organic"):
    st = R.load_style(base)
    st["arrangement"] = name
    return st


@pytest.fixture(scope="module")
def arranged(tmp_path_factory):
    out = {}
    for name in AR.NAMES:
        dst = tmp_path_factory.mktemp(name) / "pack"
        out[name] = R.run(FIXTURE, style_with(name), dst)
    return out


def arrangement_items(report):
    return {it["id"]: it for it in items(report) if it["role"] == "arrangement"}


def units_of(pack, row_id, count):
    """{n: union box of the unit's fields} in a pack, from the field names."""
    shapes = shapes_by_id(pack / "layouts.pptx", pack / "layouts.csv")[row_id]
    out = {}
    for name, sh in shapes.items():
        link = R.TAG.match(name).group(3) if R.TAG.match(name) else name
        n = AR.unit_number(link or "")
        if n and 1 <= n <= count and sh.has_text_frame:
            out.setdefault(n, []).append(R.box(sh))
    return {n: AR._union(b) for n, b in out.items()}


def plain_row(n, top=2.3, h=3.4):
    """Units of a plain row on the 12-column grid (EMU)."""
    span = 24 // n
    return [(AR.half_x(i * span), AR.inch(top), AR.span_w(span), AR.inch(h)) for i in range(n)]


def plain_column(n, x=0.6, w=7.4, top=2.4, h=0.8, step=1.1):
    return [(AR.inch(x), AR.inch(top + i * step), AR.inch(w), AR.inch(h)) for i in range(n)]


# ------------------------------------------------------------------ pure geometry
@pytest.mark.parametrize("name", AR.NAMES)
@pytest.mark.parametrize("n", [2, 3, 4, 5, 6])
def test_row_positions_no_overlap_inside_margins_in_order(name, n):
    units = plain_row(n)
    region = AR._union(units)
    region = (region[0], region[1], region[2], AR.inch(AR.GRID["content_bottom"]) - region[1])
    try:
        rects = AR.positions(name, "row", units, region)
    except AR.ArrangeError:
        assert (name == "arc" and n < 3) or (name == "deconstructed" and n not in AR.DECONSTRUCTED_ROW)
        return
    assert len(rects) == n
    assert not AR.overlapping(rects)
    lim = AR.bounds((AR.inch(AR.GRID["w"]), AR.inch(7.5)), AR._union(units))
    assert all(AR.contains(lim, r, 2) for r in rects)
    assert AR.reading_order("row", rects) == list(range(n))
    assert all(r[2] <= u[2] + 2 and r[3] <= u[3] for r, u in zip(rects, units))  # never larger


@pytest.mark.parametrize("name", AR.NAMES)
@pytest.mark.parametrize("n", [2, 3, 4, 5, 6])
def test_column_positions_no_overlap_in_order(name, n):
    units = plain_column(n)
    try:
        rects = AR.positions(name, "column", units, AR._union(units))
    except AR.ArrangeError:
        assert name == "arc" and n < 3
        return
    assert not AR.overlapping(rects)
    assert AR.reading_order("column", rects) == list(range(n))
    right = {r[0] + r[2] for r in rects}
    assert all(r[0] >= units[0][0] for r in rects) and max(right) <= units[0][0] + units[0][2]


def test_deconstructed_grid_rules():
    """Starts in half columns at most 1.5 columns from the plain grid, one span per set, never
    wider than the plain span, a gutter at least between neighbours, drops in baselines."""
    for n, table in AR.DECONSTRUCTED_ROW.items():
        base = 24 // n
        assert len({sp for _, sp, _ in table}) == 1
        for i, (start, span, drop) in enumerate(table):
            assert isinstance(start, int) and isinstance(span, int) and isinstance(drop, int)
            assert abs(start - i * base) <= AR.MAX_SHIFT_HALF
            assert span <= base and 0 <= start and start + span <= 24 and drop >= 0
        for (s0, sp0, _), (s1, _, _) in zip(table, table[1:]):
            assert s0 + sp0 <= s1
    assert max(AR.DECONSTRUCTED_COLUMN) <= AR.MAX_SHIFT_HALF
    # drops land on the baseline grid
    region = (AR.half_x(0), AR.inch(2.3), AR.span_w(24), AR.inch(4.4))
    rects = AR.positions("deconstructed", "row", plain_row(3), region)
    bl = AR.inch(AR.GRID["baseline"])
    assert all((r[1] - AR.inch(2.3)) % bl == 0 for r in rects)


def test_positions_are_deterministic():
    units = plain_row(4)
    region = (units[0][0], units[0][1], AR.span_w(24), AR.inch(4.45))
    for name in AR.NAMES:
        assert AR.positions(name, "row", units, region) == AR.positions(name, "row", list(units), region)


def test_overlap_detection():
    a = (0, 0, 100, 100)
    assert AR.intersects(a, (50, 50, 100, 100))
    assert not AR.intersects(a, (100, 0, 50, 50))  # touching edges do not overlap
    assert AR.overlapping([a, (200, 0, 10, 10), (90, 90, 20, 20)]) == [(0, 2)]


def test_reading_order_detects_swaps():
    assert AR.reading_order("row", [(0, 500, 10, 10), (20, 0, 10, 10)]) == [0, 1]  # lower but left: still first
    assert AR.reading_order("row", [(30, 0, 10, 10), (0, 0, 10, 10)]) == [1, 0]
    assert AR.reading_order("column", [(0, 30, 10, 10), (50, 0, 10, 10)]) == [1, 0]


def test_fit_stack_gives_up_slack_not_text():
    boxes = [(0, 0, 10, 100), (0, 120, 10, 300)]  # heading, gap, body with slack
    out = AR.fit_stack(boxes, [100, 120], 1000, 300)
    assert out[0][1] == 1000 and out[0][3] == 100
    assert out[1][1] + out[1][3] <= 1300 and out[1][3] >= 120 and out[1][1] >= out[0][1] + out[0][3]
    assert AR.fit_stack(boxes, [100, 300], 0, 300) is None
    # side by side (a mark beside a label) moves as one band
    band = AR.fit_stack([(0, 0, 10, 40), (20, 5, 50, 30), (0, 60, 70, 200)], [40, 30, 50], 0, 150)
    assert band[0][1] == 0 and band[1][1] == 5


def test_unknown_arrangement_refused():
    with pytest.raises(ValueError):
        AR.style_arrangement({"name": "x", "arrangement": "spiral"})
    assert AR.style_arrangement({"arrangement": {"name": "arc"}}) == "arc"
    assert AR.style_arrangement({"arrangement": None}) is None
    assert AR.parse_arrange("column:5:ordered") == ("column", 5, True)
    assert AR.parse_arrange("") is None and AR.parse_arrange("row") is None


# ------------------------------------------------------------------ on the Make_Clean fixture
def test_only_tagged_layouts_change(arranged):
    _, rows = R.read_index(FIXTURE / "layouts.csv")
    tagged = {r["id"] for r in rows if r["arrange"]}
    for name, (pack, report) in arranged.items():
        got = arrangement_items(report)
        assert set(got) == tagged | {r["id"] for r in R.read_index(FIXTURE / "visuals.csv")[1] if r["arrange"]}
        for row_id, it in got.items():
            if row_id.startswith("V-"):
                assert it["status"] == "kept"
        assert report["hooks"]["arrangement"]["applied"] is True


def test_arranged_units_keep_margins_order_and_no_overlap(arranged):
    _, rows = R.read_index(FIXTURE / "layouts.csv")
    by_id = {r["id"]: r for r in rows}
    seen = 0
    for name, (pack, report) in arranged.items():
        for row_id, it in arrangement_items(report).items():
            if it["status"] != "arranged":
                continue
            seen += 1
            kind, count, _ = AR.parse_arrange(by_id[row_id]["arrange"])
            units = units_of(pack, row_id, count)
            rects = [units[n] for n in sorted(units)]
            assert not AR.overlapping(rects), (name, row_id)
            assert AR.reading_order(kind, rects) == list(range(count)), (name, row_id)
            lim = AR.bounds((AR.inch(AR.GRID["w"]), AR.inch(7.5)), AR._union(rects))
            mx = AR.inch(AR.GRID["margin_x"])
            assert all(r[0] >= mx - R.TOL and r[0] + r[2] <= AR.inch(AR.GRID["w"]) - mx + R.TOL for r in rects)
            assert all(r[1] + r[3] <= lim[1] + lim[3] + R.TOL for r in rects)
    assert seen >= 8  # every arrangement applies somewhere in the fixture


def test_arrangement_moves_units(arranged):
    before = shapes_by_id(FIXTURE / "layouts.pptx", FIXTURE / "layouts.csv")
    pack, report = arranged["stagger"]
    assert arrangement_items(report)["L-38"]["status"] == "arranged"
    after = shapes_by_id(pack / "layouts.pptx", pack / "layouts.csv")
    tops = [R.box(after["L-38"][f"KPI {n} value"])[1] - R.box(before["L-38"][f"KPI {n} value"])[1] for n in (1, 2)]
    assert tops[1] - tops[0] >= AR.inch(0.5)
    pack, report = arranged["cascade"]
    if arrangement_items(report)["L-54"]["status"] == "arranged":
        after = shapes_by_id(pack / "layouts.pptx", pack / "layouts.csv")
        xs = [R.box(after["L-54"][f"Step {n}"])[0] for n in range(1, 6)]
        assert xs == sorted(xs) and xs[-1] > xs[0]


def test_max_content_fits_arranged_fields(arranged):
    """At max content each moved text field still fits: its new limit is within the style
    contract formula for its final box (per line x lines x 0.85), and it keeps 60 % of the old
    one. The final box and limit are those of the pack: the shape layer runs after the arrangement."""
    for name, (pack, report) in arranged.items():
        after = shapes_by_id(pack / "layouts.pptx", pack / "layouts.csv")
        final = {r["id"]: R.parse_fields(r["fields"]) for r in R.read_index(pack / "layouts.csv")[1]}
        for row_id, it in arrangement_items(report).items():
            for field, f in (it.get("fields") or {}).items():
                f = {"limit": final[row_id][field]["limit"], "old_limit": f["old_limit"]}
                sh = after[row_id][field]
                size, em, caps = R.font_info(sh, {"heading": 0.52, "body": 0.52})
                unit = size * em * (1.2 if caps else 1.0)
                li, ti, ri, bi = R.body_insets(sh)
                x, y, w, h = R.box(sh)
                lines = math.floor((h - ti - bi) / 12700 / (size * 1.2))
                per_line = math.floor((w - li - ri) / 12700 / unit)
                assert lines >= 1 and f["limit"] <= max(4, math.floor(per_line * lines * 0.85)), (name, row_id, field)
                assert f["limit"] >= math.ceil(0.6 * f["old_limit"]) or f["old_limit"] <= 4


def test_csv_limits_lowered_only(arranged):
    _, rows = R.read_index(FIXTURE / "layouts.csv")
    old = {r["id"]: R.parse_fields(r["fields"]) for r in rows}
    for pack, _ in arranged.values():
        _, new_rows = R.read_index(pack / "layouts.csv")
        for r in new_rows:
            for name, f in R.parse_fields(r["fields"]).items():
                assert set(R.parse_fields(r["fields"])) == set(old[r["id"]])
                if f["limit"] is not None:
                    assert f["limit"] <= old[r["id"]][name]["limit"]


def test_fallback_keeps_plain_grid(tmp_path):
    """When a check fails the slide is left exactly as the shape layer alone leaves it."""
    st = style_with("cascade")
    st["keep"] = 0.99  # nothing may lose any room: every arrangement that shrinks a text falls back
    _, report = R.run(FIXTURE, st, tmp_path / "a")
    plain = R.load_style("organic")
    plain["keep"] = 0.99
    _, ref = R.run(FIXTURE, plain, tmp_path / "b")
    got = arrangement_items(report)
    falls = [i for i, it in got.items() if it["status"] == "fallback"]
    assert falls and all(got[i]["reason"].startswith("plain grid kept") for i in falls)
    a = shapes_by_id(tmp_path / "a" / "layouts.pptx", tmp_path / "a" / "layouts.csv")
    b = shapes_by_id(tmp_path / "b" / "layouts.pptx", tmp_path / "b" / "layouts.csv")
    for i in falls:
        assert {n: R.box(s) for n, s in a[i].items()} == {n: R.box(s) for n, s in b[i].items()}


def test_separators_removed_only_when_arranged(arranged):
    pack, report = arranged["stagger"]
    it = arrangement_items(report)["L-54"]
    names = shapes_by_id(pack / "layouts.pptx", pack / "layouts.csv")["L-54"]
    if it["status"] == "arranged":
        assert it.get("removed") == ["decoration#1"] and "decoration#1" not in names


def test_deterministic_bytes_with_arrangement(tmp_path):
    st = style_with("deconstructed", "paper")
    a, _ = R.run(FIXTURE, st, tmp_path / "a")
    b, _ = R.run(FIXTURE, st, tmp_path / "b")
    for f in ("layouts.pptx", "layouts.csv", "restyle-report.json"):
        if f.endswith(".json"):
            ja, jb = json.loads((a / f).read_text()), json.loads((b / f).read_text())
            ja.pop("input", None), jb.pop("input", None)
            assert ja == jb
        else:
            assert (a / f).read_bytes() == (b / f).read_bytes()
    Presentation(str(a / "layouts.pptx"))


def test_no_arrangement_no_change(tmp_path):
    """A style without an arrangement gives the same pack as before this layer existed."""
    _, report = R.run(FIXTURE, R.load_style("organic"), tmp_path / "p")
    assert not arrangement_items(report)
