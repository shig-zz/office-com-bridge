from __future__ import annotations

import itertools
from types import SimpleNamespace

import pywintypes
import pytest

from office_bridge.errors import BridgeError, ErrorCode
from office_bridge.operations import excel
from office_bridge.target import AppKind


class FakeAreas:
    def __init__(self, areas: list["FakeRange"] | None = None) -> None:
        self._areas = areas or []
        self.Count = len(self._areas) if self._areas else 1

    def __call__(self, index: int) -> "FakeRange":
        return self._areas[index - 1]


class FakeHyperlink:
    def __init__(self, address: str = "") -> None:
        self.Address = address
        self.SubAddress = ""
        self.ScreenTip = ""


class FakeHyperlinks:
    def __init__(self, links: list[FakeHyperlink] | None = None) -> None:
        self._links = links or []
        self.add_calls: list[tuple[Any, str, tuple[Any, ...]]] = []

    @property
    def Count(self) -> int:
        return len(self._links)

    def __call__(self, index: int) -> FakeHyperlink:
        return self._links[index - 1]

    def Add(self, range_object: Any, address: str, *args: Any) -> FakeHyperlink:
        link = FakeHyperlink(address)
        if len(args) >= 1:
            link.SubAddress = args[0]
        if len(args) >= 2:
            link.ScreenTip = args[1]
        self.add_calls.append((range_object, address, args))
        self._links.append(link)
        return link


class FakeRange:
    def __init__(
        self,
        rows: int,
        columns: int,
        values,
        formulas=None,
        address: str = "A1:B2",
        areas: list["FakeRange"] | None = None,
        cells: list[list] | None = None,
    ) -> None:
        self.Rows = SimpleNamespace(Count=rows)
        self.Columns = SimpleNamespace(Count=columns)
        self.Count = rows * columns
        self._values = values
        self._formulas = formulas if formulas is not None else values
        self._address = address
        self.Areas = FakeAreas(areas)
        self.written = None
        self.Font = SimpleNamespace(
            Name=None,
            NameFarEast=None,
            Size=None,
            Bold=None,
            Italic=None,
            Underline=None,
            Color=None,
        )
        self.Interior = SimpleNamespace(Pattern=None, Color=None)
        self.Borders = SimpleNamespace(LineStyle=None, Color=None)
        self.NumberFormat = None
        self.HorizontalAlignment = None
        self.VerticalAlignment = None
        self.WrapText = None
        self.Hyperlinks = FakeHyperlinks()
        self.copied = False
        self.pasted_with = None
        self._cells = cells

    def Cells(self, row: int, column: int):
        if self._cells is None:
            raise AttributeError("this fake range has no per-cell grid")
        return self._cells[row - 1][column - 1]

    def Copy(self) -> None:
        self.copied = True

    def PasteSpecial(self, what: int) -> None:
        self.pasted_with = what

    @property
    def Value2(self):
        return self._values

    @Value2.setter
    def Value2(self, value) -> None:
        self.written = value
        self._values = value

    @property
    def Formula(self):
        return self._formulas

    def Address(self, row_absolute=True, column_absolute=True) -> str:
        return self._address


class FakeSheet:
    def __init__(self, ranges: dict[str, FakeRange]) -> None:
        self._ranges = ranges

    def Range(self, address: str) -> FakeRange:
        return self._ranges[address]


class FakeWindow:
    def __init__(self, selection: FakeRange, ranges: dict[str, FakeRange]) -> None:
        self.RangeSelection = selection
        self.ActiveSheet = FakeSheet(ranges)


def make_target(window: FakeWindow, document: str = "Book1.xlsx") -> SimpleNamespace:
    return SimpleNamespace(
        app=AppKind.EXCEL,
        window=window,
        document=None,
        application=SimpleNamespace(CutCopyMode=True),
        info={
            "application": "Excel",
            "document": document,
            "full_name": rf"C:\docs\{document}",
        },
    )


def test_address_normalizes_absolute_reference() -> None:
    assert excel._address(SimpleNamespace(Address="$A$1:$B$2")) == "A1:B2"


def test_read_range_returns_values_and_formulas() -> None:
    target_range = FakeRange(
        2,
        2,
        ((1.0, 2.0), (3.0, 4.0)),
        formulas=(("1", "2"), ("3", "4")),
        address="A1:B2",
    )
    target = make_target(FakeWindow(target_range, {"A1:B2": target_range}))
    result = excel.read_range(
        target, address="A1:B2", include_formulas=True, max_cells=100
    )
    assert result["values"] == [[1.0, 2.0], [3.0, 4.0]]
    assert result["formulas"] == [["1", "2"], ["3", "4"]]
    assert result["rows"] == 2
    assert result["columns"] == 2
    assert result["multi_area"] is False


def test_single_cell_is_wrapped_in_a_matrix() -> None:
    target_range = FakeRange(1, 1, "hello", address="A1")
    target = make_target(FakeWindow(target_range, {"A1": target_range}))
    result = excel.read_range(target, address="A1", include_formulas=False, max_cells=10)
    assert result["values"] == [["hello"]]


def test_read_range_rejects_large_ranges() -> None:
    target_range = FakeRange(100, 100, None)
    target = make_target(FakeWindow(target_range, {"A1:CV100": target_range}))
    with pytest.raises(BridgeError) as excinfo:
        excel.read_range(target, address="A1:CV100", include_formulas=False, max_cells=99)
    assert excinfo.value.code == ErrorCode.RANGE_TOO_LARGE


def test_selection_uses_range_selection() -> None:
    selected = FakeRange(1, 2, ((1, 2),), address="B2:C2")
    target = make_target(FakeWindow(selected, {}))
    result = excel.selection(target, limit=100)
    assert result["address"] == "B2:C2"
    assert result["values"] == [[1, 2]]
    assert result["selection_type"] == "Range"


def test_write_dry_run_does_not_require_confirmation() -> None:
    target_range = FakeRange(2, 2, None, address="D1:E2")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"D1:E2": target_range}))
    result = excel.write_range(
        target,
        address="D1:E2",
        values=[[1, 2], [3, 4]],
        dry_run=True,
        max_cells=100,
        params={},
    )
    assert result["dry_run"] is True
    assert result["written"] is False
    assert target_range.written is None


def test_write_requires_confirmation() -> None:
    target_range = FakeRange(2, 2, None, address="D1:E2")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"D1:E2": target_range}))
    with pytest.raises(BridgeError) as excinfo:
        excel.write_range(
            target,
            address="D1:E2",
            values=[[1, 2], [3, 4]],
            dry_run=False,
            max_cells=100,
            params={},
        )
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED
    assert target_range.written is None


def test_write_sets_value2_when_confirmed() -> None:
    target_range = FakeRange(2, 2, None, address="D1:E2")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"D1:E2": target_range}))
    result = excel.write_range(
        target,
        address="D1:E2",
        values=[[1, 2], [3, 4]],
        dry_run=False,
        max_cells=100,
        params={"confirm": True},
    )
    assert result["written"] is True
    assert target_range.written == ((1, 2), (3, 4))


def test_write_rejects_dimension_mismatch() -> None:
    target_range = FakeRange(3, 2, None, address="D1:E3")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"D1:E3": target_range}))
    with pytest.raises(BridgeError) as excinfo:
        excel.write_range(
            target,
            address="D1:E3",
            values=[[1, 2], [3, 4]],
            dry_run=False,
            max_cells=100,
            params={"confirm": True},
        )
    assert excinfo.value.code == ErrorCode.SIZE_MISMATCH


def test_write_rejects_formula_strings() -> None:
    target_range = FakeRange(1, 1, None, address="D1")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"D1": target_range}))
    with pytest.raises(BridgeError):
        excel.write_range(
            target,
            address="D1",
            values=[["=SUM(A1:A2)"]],
            dry_run=False,
            max_cells=100,
            params={"confirm": True},
        )


def test_write_rejects_wrong_expected_document() -> None:
    target_range = FakeRange(1, 1, None, address="D1")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"D1": target_range}))
    with pytest.raises(BridgeError) as excinfo:
        excel.write_range(
            target,
            address="D1",
            values=[[1]],
            dry_run=True,
            max_cells=100,
            params={"expected_document": "Other.xlsx"},
        )
    assert excinfo.value.code == ErrorCode.DOCUMENT_MISMATCH


def test_copy_format_dry_run_does_not_touch_ranges() -> None:
    source = FakeRange(2, 2, None, address="A1:B2")
    dest = FakeRange(2, 2, [[1, 2], [3, 4]], address="C1:D2")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"A1:B2": source, "C1:D2": dest}))
    result = excel.copy_format(
        target,
        source_address="A1:B2",
        address="C1:D2",
        dry_run=True,
        max_cells=100,
        params={},
    )
    assert result["copied"] is False
    assert source.copied is False
    assert dest.pasted_with is None


def test_copy_format_requires_confirmation() -> None:
    source = FakeRange(2, 2, None, address="A1:B2")
    dest = FakeRange(2, 2, [[1, 2], [3, 4]], address="C1:D2")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"A1:B2": source, "C1:D2": dest}))
    with pytest.raises(BridgeError) as excinfo:
        excel.copy_format(
            target,
            source_address="A1:B2",
            address="C1:D2",
            dry_run=False,
            max_cells=100,
            params={},
        )
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED
    assert dest.pasted_with is None


def test_copy_format_uses_paste_special_formats_and_clears_cut_copy_mode() -> None:
    source = FakeRange(2, 2, None, address="A1:B2")
    dest = FakeRange(2, 2, [[1, 2], [3, 4]], address="C1:D2")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"A1:B2": source, "C1:D2": dest}))
    result = excel.copy_format(
        target,
        source_address="A1:B2",
        address="C1:D2",
        dry_run=False,
        max_cells=100,
        params={"confirm": True},
    )
    assert result["copied"] is True
    assert source.copied is True
    assert dest.pasted_with == -4122  # xlPasteFormats
    assert target.application.CutCopyMode is False
    # Values are untouched by a format-only copy.
    assert dest._values == [[1, 2], [3, 4]]


def test_copy_format_rejects_dimension_mismatch() -> None:
    source = FakeRange(2, 2, None, address="A1:B2")
    dest = FakeRange(3, 2, None, address="C1:D3")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"A1:B2": source, "C1:D3": dest}))
    with pytest.raises(BridgeError) as excinfo:
        excel.copy_format(
            target,
            source_address="A1:B2",
            address="C1:D3",
            dry_run=True,
            max_cells=100,
            params={},
        )
    assert excinfo.value.code == ErrorCode.SIZE_MISMATCH


def test_set_format_requires_at_least_one_option() -> None:
    rng = FakeRange(2, 2, None, address="A1:B2")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"A1:B2": rng}))
    with pytest.raises(BridgeError) as excinfo:
        excel.set_format(
            target,
            address="A1:B2",
            dry_run=True,
            max_cells=100,
            params={},
        )
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_set_format_dry_run_reports_changes_without_applying() -> None:
    rng = FakeRange(2, 2, None, address="A1:B2")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"A1:B2": rng}))
    result = excel.set_format(
        target,
        address="A1:B2",
        dry_run=True,
        max_cells=100,
        params={"font_name": "MS Gothic", "bold": True, "halign": "center"},
    )
    assert result["applied"] is False
    assert "Font.Name='MS Gothic'" in result["changes"]
    assert "Font.Bold=True" in result["changes"]
    assert "HorizontalAlignment='center'" in result["changes"]
    # Nothing was actually applied on a dry run.
    assert rng.Font.Name is None
    assert rng.Font.Bold is None
    assert rng.HorizontalAlignment is None


def test_set_format_requires_confirmation() -> None:
    rng = FakeRange(2, 2, None, address="A1:B2")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"A1:B2": rng}))
    with pytest.raises(BridgeError) as excinfo:
        excel.set_format(
            target,
            address="A1:B2",
            dry_run=False,
            max_cells=100,
            params={"bold": True},
        )
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED
    assert rng.Font.Bold is None


def test_set_format_applies_properties_when_confirmed() -> None:
    rng = FakeRange(2, 2, None, address="A1:B2")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"A1:B2": rng}))
    result = excel.set_format(
        target,
        address="A1:B2",
        dry_run=False,
        max_cells=100,
        params={
            "confirm": True,
            "font_name": "MS Gothic",
            "font_size": 12,
            "bold": True,
            "color": "FF0000",
            "number_format": "yyyy/m/d",
            "halign": "center",
            "valign": "middle",
            "wrap": True,
            "fill": "FFFF00",
            "border": True,
            "border_color": "000000",
        },
    )
    assert result["applied"] is True
    assert rng.Font.Name == "MS Gothic"
    assert rng.Font.Size == 12.0
    assert rng.Font.Bold is True
    # FF0000 (red) is stored by Excel as BGR 0x0000FF.
    assert rng.Font.Color == 0x0000FF
    assert rng.NumberFormat == "yyyy/m/d"
    assert rng.HorizontalAlignment == -4108  # xlHAlignCenter
    assert rng.VerticalAlignment == -4108  # xlVAlignCenter
    assert rng.WrapText is True
    assert rng.Interior.Pattern == 1  # xlPatternSolid
    assert rng.Interior.Color == 0x00FFFF  # FFFF00 (yellow) as BGR
    assert rng.Borders.LineStyle == 1  # xlContinuous
    assert rng.Borders.Color == 0x000000


def test_set_format_rejects_font_name_far_east_with_guidance() -> None:
    rng = FakeRange(1, 1, None, address="A1")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"A1": rng}))
    with pytest.raises(BridgeError) as excinfo:
        excel.set_format(
            target,
            address="A1",
            dry_run=True,
            max_cells=100,
            params={"font_name_far_east": "游ゴシック"},
        )
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST
    assert "pan-CJK" in excinfo.value.message
    assert rng.Font.Name is None


def test_set_format_rejects_invalid_color() -> None:
    rng = FakeRange(1, 1, None, address="A1")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"A1": rng}))
    with pytest.raises(BridgeError) as excinfo:
        excel.set_format(
            target,
            address="A1",
            dry_run=True,
            max_cells=100,
            params={"color": "nope"},
        )
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_set_format_rejects_invalid_alignment() -> None:
    rng = FakeRange(1, 1, None, address="A1")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"A1": rng}))
    with pytest.raises(BridgeError) as excinfo:
        excel.set_format(
            target,
            address="A1",
            dry_run=True,
            max_cells=100,
            params={"halign": "diagonal"},
        )
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_add_hyperlink_dry_run_does_not_add() -> None:
    rng = FakeRange(1, 1, "https://example.com", address="E35")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"E35": rng}))
    result = excel.add_hyperlink(
        target,
        address="E35",
        dry_run=True,
        max_cells=100,
        params={"url": "https://example.com"},
    )
    assert result["dry_run"] is True
    assert result["applied"] is False
    assert result["action"] == "add"
    assert result["existing_links"] == 0
    assert rng.Hyperlinks.add_calls == []


def test_add_hyperlink_requires_confirmation() -> None:
    rng = FakeRange(1, 1, "https://example.com", address="E35")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"E35": rng}))
    with pytest.raises(BridgeError) as excinfo:
        excel.add_hyperlink(
            target,
            address="E35",
            dry_run=False,
            max_cells=100,
            params={"url": "https://example.com"},
        )
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED
    assert rng.Hyperlinks.add_calls == []


def test_add_hyperlink_adds_link_when_confirmed() -> None:
    rng = FakeRange(1, 1, "https://example.com", address="E35")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"E35": rng}))
    result = excel.add_hyperlink(
        target,
        address="E35",
        dry_run=False,
        max_cells=100,
        params={"confirm": True, "url": "https://example.com"},
    )
    assert result["applied"] is True
    assert result["action"] == "add"
    # Hyperlinks.Add receives the range object itself as its first argument.
    link_range, link_url, extra = rng.Hyperlinks.add_calls[0]
    assert link_range is rng
    assert link_url == "https://example.com"
    assert extra == ()
    assert rng.Hyperlinks.Count == 1
    assert rng.Hyperlinks(1).Address == "https://example.com"


def test_add_hyperlink_passes_sub_address_and_screen_tip() -> None:
    rng = FakeRange(1, 1, "https://example.com", address="E35")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"E35": rng}))
    result = excel.add_hyperlink(
        target,
        address="E35",
        dry_run=False,
        max_cells=100,
        params={
            "confirm": True,
            "url": "https://example.com",
            "sub_address": "Sheet2!A1",
            "screen_tip": "公式ページ",
        },
    )
    assert result["applied"] is True
    _, _, extra = rng.Hyperlinks.add_calls[0]
    assert extra == ("Sheet2!A1", "公式ページ")
    link = rng.Hyperlinks(1)
    assert link.SubAddress == "Sheet2!A1"
    assert link.ScreenTip == "公式ページ"


def test_add_hyperlink_updates_existing_link_instead_of_adding() -> None:
    existing = FakeHyperlink("https://old.example.com")
    existing.ScreenTip = "古いヒント"
    rng = FakeRange(1, 1, "https://old.example.com", address="E35")
    rng.Hyperlinks = FakeHyperlinks([existing])
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"E35": rng}))
    result = excel.add_hyperlink(
        target,
        address="E35",
        dry_run=False,
        max_cells=100,
        params={
            "confirm": True,
            "url": "https://new.example.com",
            "screen_tip": "新しいヒント",
        },
    )
    assert result["applied"] is True
    assert result["action"] == "update"
    assert result["existing_links"] == 1
    assert rng.Hyperlinks.add_calls == []
    assert rng.Hyperlinks.Count == 1
    assert existing.Address == "https://new.example.com"
    assert existing.ScreenTip == "新しいヒント"


def test_add_hyperlink_dry_run_reports_update_for_existing_link() -> None:
    rng = FakeRange(1, 1, "https://old.example.com", address="E35")
    rng.Hyperlinks = FakeHyperlinks([FakeHyperlink("https://old.example.com")])
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"E35": rng}))
    result = excel.add_hyperlink(
        target,
        address="E35",
        dry_run=True,
        max_cells=100,
        params={"url": "https://new.example.com"},
    )
    assert result["applied"] is False
    assert result["action"] == "update"
    assert rng.Hyperlinks(1).Address == "https://old.example.com"


def test_add_hyperlink_requires_url() -> None:
    rng = FakeRange(1, 1, "https://example.com", address="E35")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"E35": rng}))
    with pytest.raises(BridgeError) as excinfo:
        excel.add_hyperlink(target, address="E35", dry_run=True, max_cells=100, params={})
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_add_hyperlink_rejects_formula_addresses() -> None:
    rng = FakeRange(1, 1, "https://example.com", address="E35")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"E35": rng}))
    with pytest.raises(BridgeError) as excinfo:
        excel.add_hyperlink(
            target,
            address="E35",
            dry_run=True,
            max_cells=100,
            params={"url": "https://example.com", "sub_address": "=A1"},
        )
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_add_hyperlink_rejects_invalid_screen_tip() -> None:
    rng = FakeRange(1, 1, "https://example.com", address="E35")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"E35": rng}))
    with pytest.raises(BridgeError) as excinfo:
        excel.add_hyperlink(
            target,
            address="E35",
            dry_run=True,
            max_cells=100,
            params={"url": "https://example.com", "screen_tip": 42},
        )
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_add_hyperlink_rejects_wrong_expected_document() -> None:
    rng = FakeRange(1, 1, "https://example.com", address="E35")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"E35": rng}))
    with pytest.raises(BridgeError) as excinfo:
        excel.add_hyperlink(
            target,
            address="E35",
            dry_run=True,
            max_cells=100,
            params={"url": "https://example.com", "expected_document": "Other.xlsx"},
        )
    assert excinfo.value.code == ErrorCode.DOCUMENT_MISMATCH


def test_column_letter_roundtrip() -> None:
    assert excel._column_letter(1) == "A"
    assert excel._column_letter(26) == "Z"
    assert excel._column_letter(27) == "AA"
    assert excel._column_letter(702) == "ZZ"
    assert excel._column_letter(703) == "AAA"
    assert excel._letter_column("C") == 3
    assert excel._letter_column("AA") == 27


def test_compute_order_single_key_ascending_is_stable() -> None:
    values = [[2, "b"], [1, "z"], [2, "a"], [1, "y"]]
    assert excel._compute_order(values, [(1, True)]) == [1, 3, 0, 2]


def test_compute_order_multi_key_breaks_ties() -> None:
    values = [["a", "2"], ["b", "1"], ["a", "1"], ["b", "2"]]
    assert excel._compute_order(values, [(1, True), (2, True)]) == [2, 0, 1, 3]


def test_compute_order_descending() -> None:
    values = [[2], [1], [3]]
    assert excel._compute_order(values, [(1, False)]) == [2, 0, 1]


def test_compute_order_mixed_types_order_blanks_numbers_text() -> None:
    values = [["x"], [3.0], [None], [1], ["a"]]
    assert excel._compute_order(values, [(1, True)]) == [2, 3, 1, 4, 0]


def test_moves_for_noop_when_sorted() -> None:
    assert excel._moves_for([0, 1, 2, 3]) == []


def test_moves_for_lifts_rows_from_below() -> None:
    # target order [2, 4, 3, 5, 1] as 0-based original indices
    moves = excel._moves_for([1, 3, 2, 4, 0])
    assert all(source > destination for destination, source in moves)
    # replaying the moves reproduces the target order
    current = list(range(5))
    for destination, source in moves:
        lifted = current[source - 1]
        current = (
            current[: destination - 1]
            + [lifted]
            + current[destination - 1 : source - 1]
            + current[source:]
        )
    assert current == [1, 3, 2, 4, 0]


def test_parse_sort_key_accepts_offset_and_letter() -> None:
    assert excel._parse_sort_key({"column": 2}, 3, 1) == (2, True)
    assert excel._parse_sort_key({"column": "C"}, 5, 1) == (3, True)
    assert excel._parse_sort_key({"column": "B", "order": "descending"}, 2, 1) == (
        2,
        False,
    )
    # a letter is resolved relative to the range's left edge
    assert excel._parse_sort_key({"column": "C"}, 3, 2) == (2, True)


def test_parse_sort_key_rejects_out_of_range_and_bad_order() -> None:
    for bad in ({"column": 0}, {"column": 4}, {"column": "D"}, {"column": "AA"}, {"column": 1.5}):
        with pytest.raises(BridgeError):
            excel._parse_sort_key(bad, 3, 1)
    with pytest.raises(BridgeError):
        excel._parse_sort_key({"column": 1, "order": "sideways"}, 3, 1)
    with pytest.raises(BridgeError):
        excel._parse_sort_key("column 1", 3, 1)


class FakeSortRow:
    def __init__(self, sheet: "FakeSortSheet", row: int, count: int = 1) -> None:
        self._sheet = sheet
        self._row = row
        self._count = count

    @property
    def Row(self) -> int:
        return self._row

    @property
    def Rows(self) -> SimpleNamespace:
        return SimpleNamespace(Count=self._count)

    def Copy(self) -> None:
        # whole-row copy: the entire row goes on the clipboard
        self._sheet.copies += 1
        self._sheet.clipboard = list(self._sheet.grid[self._row - 1])
        self._sheet.clip_left = 0

    def Insert(self, shift: int = -4121) -> None:
        for offset in range(self._count):
            self._sheet._insert_row(self._row + offset)

    def Delete(self, shift: int = -4121) -> None:
        for offset in reversed(range(self._count)):
            self._sheet._delete_row(self._row + offset)

    @property
    def RowHeight(self) -> float | None:
        if self._row in self._sheet.unknown_height_rows:
            return None
        return self._sheet.heights[self._row - 1]

    @RowHeight.setter
    def RowHeight(self, value: float) -> None:
        # assigning None would raise in the real COM layer
        if value is None:
            raise TypeError("float() argument must be a string or a real number, not 'NoneType'")
        self._sheet.height_sets.append((self._row, value))
        self._sheet.heights[self._row - 1] = float(value)


_ROW_ID_COUNTER = itertools.count(1)


class FakeSheetRange:
    """A live view over the fake grid, mirroring the members excel.sort_range uses.

    Like the installed Excel build, the range keeps track of the rows that existed when
    it was created: rows deleted afterwards drop out of the view, and rows inserted
    afterwards are not part of it. Re-resolving the address yields a full fresh view.
    """

    def __init__(self, sheet: "FakeSortSheet", top: int, left: int, nrows: int, ncols: int) -> None:
        self._sheet = sheet
        self.top = top
        self.left = left
        self.nrows = nrows
        self.ncols = ncols
        self._row_ids = list(sheet.row_ids[top - 1 : top - 1 + nrows])

    @property
    def Rows(self):
        return SimpleNamespace(Count=self.nrows)

    @property
    def Columns(self):
        sheet_range = self

        class _Columns:
            @property
            def Count(self) -> int:
                return sheet_range.ncols

            def __call__(self, index: int):
                return SimpleNamespace(Column=sheet_range.left + index - 1)

        return _Columns()

    @property
    def Row(self) -> int:
        return self.top

    @property
    def Value2(self):
        sheet = self._sheet
        rows = []
        for row_id in self._row_ids:
            if row_id not in sheet.row_ids:
                continue  # the row was deleted after this range was created
            index = sheet.row_ids.index(row_id)
            rows.append(tuple(sheet.grid[index][self.left - 1 : self.left - 1 + self.ncols]))
        return tuple(rows)

    @property
    def RowHeight(self) -> None:
        # In the installed build a multi-row Range.RowHeight is unreliable and can
        # report None (observed on the main sheet). sort_range must read heights
        # row by row through sheet.Rows(n).RowHeight instead.
        return None

    @property
    def Parent(self) -> "FakeSortSheet":
        return self._sheet

    def Address(self, row_absolute: bool = True, column_absolute: bool = True) -> str:
        end_row = self.top + self.nrows - 1
        end_col = excel._column_letter(self.left + self.ncols - 1)
        prefix = "$" if row_absolute else ""
        col_prefix = "$" if column_absolute else ""
        start = f"{prefix}{col_prefix}{excel._column_letter(self.left)}{self.top}"
        end = f"{prefix}{col_prefix}{end_col}{end_row}"
        return f"{start}:{end}" if (self.nrows, self.ncols) != (1, 1) else start


def _parse_a1(address: str) -> tuple[int, int, int, int]:
    text = address.upper().replace("$", "")
    if ":" in text:
        start, end = text.split(":", 1)
    else:
        start = end = text

    def cell(ref: str) -> tuple[int, int]:
        letters = 0
        while letters < len(ref) and ref[letters].isalpha():
            letters += 1
        return int(ref[letters:]), excel._letter_column(ref[:letters])

    top, left = cell(start)
    bottom, right = cell(end)
    return top, left, bottom - top + 1, right - left + 1


class FakeSortSheet:
    """In-memory sheet modelling the row-move choreography of excel.sort_range.

    Copy puts the row's cells on the clipboard; Rows(n).Insert pastes the clipboard
    into the inserted row (the behaviour of the installed Excel build); setting the
    application's CutCopyMode to False clears the clipboard.
    """

    def __init__(self, grid: list[list], heights: list[float] | None = None) -> None:
        self.grid = [list(row) for row in grid]
        self.row_ids = [next(_ROW_ID_COUNTER) for _ in self.grid]
        width = len(self.grid[0]) if self.grid else 0
        self.heights = list(heights) if heights else [24.0] * len(self.grid)
        self._width = width
        self.Name = "Sheet1"
        self.clipboard: list | None = None
        self.clip_left = 0
        self.copies = 0
        self.inserts = 0
        self.deletes = 0
        self.height_sets: list[tuple[int, float]] = []
        # rows whose Rows(n).RowHeight reports None (multi-row reads do on some
        # sheets of the installed build); per-row reads normally still work
        self.unknown_height_rows: set[int] = set()

    @property
    def UsedRange(self) -> SimpleNamespace:
        # Inserted blank rows count as used on the installed build, so the
        # used range spans the whole grid.
        return SimpleNamespace(
            Row=1,
            Rows=SimpleNamespace(Count=len(self.grid)),
            Column=1,
            Columns=SimpleNamespace(Count=self._width),
        )

    def Range(self, address: str) -> FakeSheetRange:
        top, left, nrows, ncols = _parse_a1(address)
        return FakeSheetRange(self, top, left, nrows, ncols)

    def Rows(self, row: int | str) -> FakeSortRow:
        if isinstance(row, str) and ":" in row:
            first, last = (int(part) for part in row.split(":"))
            return FakeSortRow(self, first, last - first + 1)
        return FakeSortRow(self, int(row))

    def _insert_row(self, row: int) -> None:
        self.inserts += 1
        self.grid.insert(row - 1, [None] * self._width)
        self.row_ids.insert(row - 1, next(_ROW_ID_COUNTER))
        self.heights.insert(row - 1, self.heights[row - 1] if row <= len(self.heights) else 24.0)
        if self.clipboard is not None:
            target = self.grid[row - 1]
            for offset, value in enumerate(self.clipboard):
                target[self.clip_left + offset] = value

    def _delete_row(self, row: int) -> None:
        self.deletes += 1
        del self.grid[row - 1]
        del self.row_ids[row - 1]
        del self.heights[row - 1]


def make_sort_sheet() -> FakeSortSheet:
    # row 1 / row 7 are sentinels that must never be disturbed
    return FakeSortSheet(
        [
            ["TOP", "sentinel"],
            ["r5", "e"],
            ["r1", "a"],
            ["r4", "d"],
            ["r2", "b"],
            ["r3", "c"],
            ["BOTTOM", "sentinel"],
        ]
    )


def make_sort_target(sheet: FakeSortSheet) -> SimpleNamespace:
    application = SimpleNamespace(CutCopyMode=False)
    sheet.application = application
    window = SimpleNamespace(
        RangeSelection=FakeRange(1, 1, None),
        ActiveSheet=sheet,
    )
    target = make_target(window)
    target.application = application
    return target


def test_sort_range_dry_run_reports_plan_without_changes() -> None:
    sheet = make_sort_sheet()
    target = make_sort_target(sheet)
    before = [list(row) for row in sheet.grid]
    result = excel.sort_range(
        target,
        address="A2:B6",
        keys=[{"column": 1}],
        dry_run=True,
        max_cells=100,
        params={},
    )
    assert result["dry_run"] is True
    assert result["applied"] is False
    assert result["rows"] == 5
    assert result["columns"] == 2
    assert result["keys"] == [{"column": 1, "order": "ascending"}]
    # original data rows (1-based within the range) in the sorted order:
    # r1 was row 2, r2 row 4, r3 row 5, r4 row 3, r5 row 1
    assert result["row_order"] == [2, 4, 5, 3, 1]
    assert result["moves"] == 4
    assert [list(row) for row in sheet.grid] == before
    assert sheet.copies == 0


def test_sort_range_requires_confirmation() -> None:
    sheet = make_sort_sheet()
    target = make_sort_target(sheet)
    with pytest.raises(BridgeError) as excinfo:
        excel.sort_range(
            target,
            address="A2:B6",
            keys=[{"column": 1}],
            dry_run=False,
            max_cells=100,
            params={},
        )
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED
    assert sheet.copies == 0


def test_sort_range_moves_rows_and_verifies() -> None:
    sheet = make_sort_sheet()
    target = make_sort_target(sheet)
    result = excel.sort_range(
        target,
        address="A2:B6",
        keys=[{"column": "A"}],
        dry_run=False,
        max_cells=100,
        params={"confirm": True},
    )
    assert result["applied"] is True
    assert result["verified"] is True
    assert sheet.grid[0] == ["TOP", "sentinel"]
    assert sheet.grid[1:6] == [
        ["r1", "a"],
        ["r2", "b"],
        ["r3", "c"],
        ["r4", "d"],
        ["r5", "e"],
    ]
    assert sheet.grid[6] == ["BOTTOM", "sentinel"]
    # one copy/insert/delete per move
    assert sheet.copies == sheet.inserts == sheet.deletes == 4
    assert target.application.CutCopyMode is False


def test_sort_range_multi_key_and_descending() -> None:
    sheet = FakeSortSheet(
        [
            ["a", "2"],
            ["b", "1"],
            ["a", "1"],
            ["b", "2"],
        ]
    )
    target = make_sort_target(sheet)
    result = excel.sort_range(
        target,
        address="A1:B4",
        keys=[{"column": 1}, {"column": "B", "order": "descending"}],
        dry_run=False,
        max_cells=100,
        params={"confirm": True},
    )
    assert result["verified"] is True
    # a(2), a(1), b(2), b(1): primary asc, tie broken desc
    assert sheet.grid == [["a", "2"], ["a", "1"], ["b", "2"], ["b", "1"]]


def test_sort_range_follows_non_uniform_row_heights() -> None:
    sheet = make_sort_sheet()
    # the r5 row (data row 2) has a distinct height that must follow the row
    sheet.heights[1] = 42.0
    target = make_sort_target(sheet)
    result = excel.sort_range(
        target,
        address="A2:B6",
        keys=[{"column": 1}],
        dry_run=False,
        max_cells=100,
        params={"confirm": True},
    )
    assert result["verified"] is True
    # r5 ends up last, so the 42pt height must land on the last data row (row 6)
    assert sheet.heights[1:6] == [24.0, 24.0, 24.0, 24.0, 42.0]
    assert sheet.height_sets != []


def test_sort_range_uniform_heights_need_no_writes() -> None:
    sheet = make_sort_sheet()
    target = make_sort_target(sheet)
    excel.sort_range(
        target,
        address="A2:B6",
        keys=[{"column": 1}],
        dry_run=False,
        max_cells=100,
        params={"confirm": True},
    )
    assert sheet.height_sets == []


def test_sort_range_tolerates_unknown_row_heights() -> None:
    # On some sheets of the installed build even Rows(n).RowHeight can report
    # None. The sort must still complete and must not write None back.
    sheet = make_sort_sheet()
    sheet.unknown_height_rows = {2, 3, 4, 5, 6}
    target = make_sort_target(sheet)
    result = excel.sort_range(
        target,
        address="A2:B6",
        keys=[{"column": 1}],
        dry_run=False,
        max_cells=100,
        params={"confirm": True},
    )
    assert result["verified"] is True
    assert sheet.grid[1:6] == [
        ["r1", "a"],
        ["r2", "b"],
        ["r3", "c"],
        ["r4", "d"],
        ["r5", "e"],
    ]
    assert sheet.height_sets == []


def test_sort_range_rejects_bad_keys_parameter() -> None:
    sheet = make_sort_sheet()
    target = make_sort_target(sheet)
    for bad in (None, "C", [], [1, 2, 3, 4]):
        with pytest.raises(BridgeError) as excinfo:
            excel.sort_range(
                target,
                address="A2:B6",
                keys=bad,
                dry_run=True,
                max_cells=100,
                params={},
            )
        assert excinfo.value.code == ErrorCode.INVALID_REQUEST


# ---------------------------------------------------------------------------
# Phase 3A: structural writes, format reads, copy-format borders


def test_rows_insert_dry_run_and_apply() -> None:
    sheet = make_sort_sheet()
    target = make_sort_target(sheet)
    dry = excel.insert_or_delete(
        target,
        dimension="rows",
        spec="2:4",
        shift=None,
        mode="insert",
        dry_run=True,
        params={},
    )
    assert dry["first"] == 2
    assert dry["count"] == 3
    assert dry["shift"] == "down"
    assert dry["applied"] is False
    assert len(sheet.grid) == 7

    result = excel.insert_or_delete(
        target,
        dimension="rows",
        spec="2:4",
        shift=None,
        mode="insert",
        dry_run=False,
        params={"confirm": True},
    )
    assert result["applied"] is True
    assert result["verified"] is True
    assert result["last_used_after"] == 10
    assert "undo_hint" in result
    assert len(sheet.grid) == 10


def test_rows_insert_below_used_range_still_verifies() -> None:
    # Inserting just below the used area extends the used range to the
    # inserted rows themselves; the verification must expect that instead of
    # flagging the growth as unverified (found in the live Excel smoke).
    sheet = make_sort_sheet()
    target = make_sort_target(sheet)
    result = excel.insert_or_delete(
        target,
        dimension="rows",
        spec="8:9",
        shift=None,
        mode="insert",
        dry_run=False,
        params={"confirm": True},
    )
    assert result["first"] == 8
    assert result["last_used_after"] == 9
    assert result["verified"] is True


def test_rows_insert_requires_confirmation() -> None:
    sheet = make_sort_sheet()
    target = make_sort_target(sheet)
    with pytest.raises(BridgeError) as excinfo:
        excel.insert_or_delete(
            target,
            dimension="rows",
            spec=2,
            shift=None,
            mode="insert",
            dry_run=False,
            params={},
        )
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED
    assert sheet.inserts == 0


def test_rows_insert_rejects_bad_specs() -> None:
    sheet = make_sort_sheet()
    target = make_sort_target(sheet)
    for bad in (0, "4:2", "x", True, 2.5):
        with pytest.raises(BridgeError) as excinfo:
            excel.insert_or_delete(
                target,
                dimension="rows",
                spec=bad,
                shift=None,
                mode="insert",
                dry_run=True,
                params={},
            )
        assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_rows_delete_verifies_and_shrinks_grid() -> None:
    sheet = make_sort_sheet()
    target = make_sort_target(sheet)
    result = excel.insert_or_delete(
        target,
        dimension="rows",
        spec=6,
        shift=None,
        mode="delete",
        dry_run=False,
        params={"confirm": True},
    )
    assert result["applied"] is True
    assert result["verified"] is True
    assert result["last_used_after"] == 6
    assert len(sheet.grid) == 6
    assert sheet.grid[5][0] == "BOTTOM"


def test_rows_delete_rejects_bad_shift() -> None:
    sheet = make_sort_sheet()
    target = make_sort_target(sheet)
    with pytest.raises(BridgeError) as excinfo:
        excel.insert_or_delete(
            target,
            dimension="rows",
            spec=2,
            shift="sideways",
            mode="delete",
            dry_run=True,
            params={},
        )
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


class FakeColumnBlock:
    def __init__(self, sheet: "FakeColumnSheet", first: int, count: int) -> None:
        self._sheet = sheet
        self._first = first
        self._count = count

    @property
    def Column(self) -> int:
        return self._first

    @property
    def Columns(self) -> SimpleNamespace:
        return SimpleNamespace(Count=self._count)

    def Insert(self, shift: int = -4161) -> None:
        self._sheet.width += self._count

    def Delete(self, shift: int = -4121) -> None:
        self._sheet.width -= self._count


class FakeColumnSheet:
    def __init__(self, width: int = 5) -> None:
        self.width = width
        self.Name = "Columns1"

    @property
    def UsedRange(self) -> SimpleNamespace:
        return SimpleNamespace(
            Row=1,
            Rows=SimpleNamespace(Count=3),
            Column=1,
            Columns=SimpleNamespace(Count=self.width),
        )

    def Columns(self, spec: int | str) -> FakeColumnBlock:
        if isinstance(spec, str):
            if ":" in spec:
                first, last = (excel._letter_column(part) for part in spec.split(":"))
            else:
                first = last = excel._letter_column(spec)
            return FakeColumnBlock(self, first, last - first + 1)
        return FakeColumnBlock(self, int(spec), 1)


def make_column_target(sheet: FakeColumnSheet) -> SimpleNamespace:
    window = SimpleNamespace(RangeSelection=None, ActiveSheet=sheet)
    return make_target(window)


def test_columns_insert_by_letter_and_delete_by_range() -> None:
    sheet = FakeColumnSheet(width=5)
    target = make_column_target(sheet)
    result = excel.insert_or_delete(
        target,
        dimension="columns",
        spec="C",
        shift=None,
        mode="insert",
        dry_run=False,
        params={"confirm": True},
    )
    assert result["first"] == 3
    assert result["shift"] == "right"
    assert result["verified"] is True
    assert sheet.width == 6

    deleted = excel.insert_or_delete(
        target,
        dimension="columns",
        spec="A:C",
        shift=None,
        mode="delete",
        dry_run=False,
        params={"confirm": True},
    )
    assert deleted["count"] == 3
    assert deleted["verified"] is True
    assert sheet.width == 3


def test_move_row_up_applies_and_verifies() -> None:
    sheet = FakeSortSheet([["a", 1], ["b", 2], ["c", 3], ["d", 4]])
    target = make_sort_target(sheet)
    result = excel.move_rows(
        target, row=4, to=2, sheet_name=None, dry_run=False, params={"confirm": True}
    )
    assert result["applied"] is True
    assert result["verified"] is True
    assert result["direction"] == "up"
    assert result["completed_steps"] == ["copy", "insert", "clear_clipboard", "delete_original"]
    assert sheet.grid == [["a", 1], ["d", 4], ["b", 2], ["c", 3]]


def test_move_row_down_applies_and_verifies() -> None:
    sheet = FakeSortSheet([["a", 1], ["b", 2], ["c", 3], ["d", 4]])
    target = make_sort_target(sheet)
    result = excel.move_rows(
        target, row=2, to=4, sheet_name=None, dry_run=False, params={"confirm": True}
    )
    assert result["applied"] is True
    assert result["verified"] is True
    assert result["direction"] == "down"
    assert sheet.grid == [["a", 1], ["c", 3], ["d", 4], ["b", 2]]
    assert "undo_note" in result


def test_move_row_requires_confirmation() -> None:
    sheet = make_sort_sheet()
    target = make_sort_target(sheet)
    with pytest.raises(BridgeError) as excinfo:
        excel.move_rows(
            target, row=2, to=5, sheet_name=None, dry_run=False, params={}
        )
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED
    assert sheet.copies == 0


def test_move_row_dry_run_reports_plan() -> None:
    sheet = make_sort_sheet()
    target = make_sort_target(sheet)
    result = excel.move_rows(
        target, row=6, to=2, sheet_name=None, dry_run=True, params={}
    )
    assert result["direction"] == "up"
    assert result["steps"] == ["copy", "insert", "clear_clipboard", "delete_original"]
    assert sheet.copies == 0


def test_move_row_rejects_out_of_range() -> None:
    sheet = make_sort_sheet()
    target = make_sort_target(sheet)
    with pytest.raises(BridgeError) as excinfo:
        excel.move_rows(
            target, row=99, to=2, sheet_name=None, dry_run=True, params={}
        )
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST
    with pytest.raises(BridgeError) as same_exc:
        excel.move_rows(
            target, row=2, to=2, sheet_name=None, dry_run=True, params={}
        )
    assert same_exc.value.code == ErrorCode.INVALID_REQUEST


def test_move_row_partial_failure_reports_completed_steps(monkeypatch) -> None:
    sheet = make_sort_sheet()
    target = make_sort_target(sheet)

    def boom(self, shift: int = -4121) -> None:  # noqa: ANN001
        raise pywintypes.com_error(-2147352567, "boom", None, None, 0, None)

    monkeypatch.setattr(FakeSortRow, "Delete", boom)
    with pytest.raises(BridgeError) as excinfo:
        excel.move_rows(
            target, row=6, to=2, sheet_name=None, dry_run=False, params={"confirm": True}
        )
    assert excinfo.value.code == ErrorCode.RANGE_INVALID
    assert excinfo.value.details["completed_steps"] == [
        "copy",
        "insert",
        "clear_clipboard",
    ]


class FakeBorder:
    def __init__(self, style: int = -4142, weight: int | None = None, color: int | None = None) -> None:
        self.LineStyle = style
        self.Weight = weight
        self.Color = color


class FakeBorderSet:
    def __init__(self, borders: dict[int, FakeBorder] | None = None) -> None:
        self._borders = dict(borders or {})

    def __call__(self, position: int) -> FakeBorder:
        if position not in self._borders:
            self._borders[position] = FakeBorder()
        return self._borders[position]


class _FormatColumns:
    def __init__(self, count: int, left: int) -> None:
        self.Count = count
        self._left = left

    def __call__(self, index: int) -> SimpleNamespace:
        return SimpleNamespace(Column=self._left + index - 1)


def make_format_cell(
    fill: int | None = None,
    pattern: int = -4142,
    font_color: int | None = None,
    borders: dict[int, FakeBorder] | None = None,
) -> SimpleNamespace:
    return SimpleNamespace(
        Interior=SimpleNamespace(Pattern=pattern, Color=fill),
        Font=SimpleNamespace(Name="Meiryo", Size=12.0, Bold=True, Color=font_color),
        NumberFormat="0.0%",
        HorizontalAlignment=-4108,
        VerticalAlignment=-4108,
        WrapText=False,
        Borders=FakeBorderSet(borders),
    )


class FakeFormatRange:
    def __init__(self, cells: list[list], top: int, left: int, address: str) -> None:
        self._cells = cells
        self.Rows = SimpleNamespace(Count=len(cells))
        self.Columns = _FormatColumns(len(cells[0]), left)
        self._row = top
        self._address = address

    @property
    def Row(self) -> int:
        return self._row

    def Cells(self, row: int, column: int):
        return self._cells[row - 1][column - 1]

    def Address(self, row_absolute: bool = True, column_absolute: bool = True) -> str:
        return self._address


def make_format_target(cells: list[list], top: int = 2, left: int = 2, address: str = "B2") -> SimpleNamespace:
    fmt_range = FakeFormatRange(cells, top=top, left=left, address=address)
    window = FakeWindow(FakeRange(1, 1, None), {address: fmt_range})
    return make_target(window)


def test_read_format_reads_cells_with_hex_colors() -> None:
    plain = make_format_cell()
    filled = make_format_cell(fill=0x0000FA, pattern=1, font_color=255)
    target = make_format_target([[filled, plain], [plain, plain]], address="B2:C3")
    result = excel.read_format(target, address="B2:C3", max_cells=100)
    assert result["address"] == "B2:C3"
    assert result["cell_count"] == 4
    assert result["borders"] is False
    first = result["cells"][0]
    assert first["cell"] == "B2"
    assert first["fill"] == "FA0000"  # stored BGR 0000FA -> RRGGBB FA0000
    assert first["font"] == {"name": "Meiryo", "size": 12.0, "bold": True, "color": "FF0000"}
    assert first["number_format"] == "0.0%"
    assert "borders" not in first
    assert result["cells"][1]["cell"] == "C2"
    assert result["cells"][1]["fill"] is None
    assert result["cells"][3]["cell"] == "C3"


def test_read_format_includes_borders_when_requested() -> None:
    cell = make_format_cell(borders={9: FakeBorder(style=1, weight=4, color=0x303030)})
    target = make_format_target([[cell]], address="B2")
    result = excel.read_format(target, address="B2", max_cells=10, include_borders=True)
    borders = result["cells"][0]["borders"]
    assert borders["bottom"]["style"] == "continuous"
    assert borders["bottom"]["weight"] == 4
    assert borders["bottom"]["color"] == "303030"
    assert borders["top"]["style"] == "none"
    assert "weight" not in borders["top"]


def test_read_format_respects_cell_limit() -> None:
    cell = make_format_cell()
    target = make_format_target([[cell, cell], [cell, cell]], address="B2:C3")
    with pytest.raises(BridgeError) as excinfo:
        excel.read_format(target, address="B2:C3", max_cells=3)
    assert excinfo.value.code == ErrorCode.RANGE_TOO_LARGE


def test_copy_format_default_reports_borders_note() -> None:
    source = FakeRange(1, 1, None, address="A1")
    dest = FakeRange(1, 1, None, address="C1")
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"A1": source, "C1": dest}))
    result = excel.copy_format(
        target,
        source_address="A1",
        address="C1",
        dry_run=False,
        max_cells=10,
        params={"confirm": True},
    )
    assert result["copied"] is True
    assert result["borders_applied"] is False
    assert "borders" in result["borders_note"]


def test_copy_format_borders_restores_per_cell() -> None:
    source_cell = SimpleNamespace(
        Borders=FakeBorderSet(
            {9: FakeBorder(style=1, weight=4, color=0x303030)}
        )
    )
    dest_cell = SimpleNamespace(Borders=FakeBorderSet())
    source = FakeRange(1, 1, None, address="A1", cells=[[source_cell]])
    dest = FakeRange(1, 1, None, address="C1", cells=[[dest_cell]])
    target = make_target(FakeWindow(FakeRange(1, 1, None), {"A1": source, "C1": dest}))
    result = excel.copy_format(
        target,
        source_address="A1",
        address="C1",
        dry_run=False,
        max_cells=10,
        params={"confirm": True},
        include_borders=True,
    )
    assert result["borders_applied"] is True
    bottom = dest_cell.Borders(9)
    assert bottom.LineStyle == 1
    assert bottom.Weight == 4
    assert bottom.Color == 0x303030
