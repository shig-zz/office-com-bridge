from __future__ import annotations

from types import SimpleNamespace

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
