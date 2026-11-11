from __future__ import annotations

from types import SimpleNamespace

import pytest

from office_bridge.errors import BridgeError, ErrorCode
from office_bridge.operations import powerpoint
from office_bridge.target import AppKind


class _Caller:
    """Callable wrapper: SimpleNamespace does not honour a __call__ attribute."""

    def __init__(self, fn) -> None:
        self._fn = fn

    def __call__(self, *args):
        return self._fn(*args)


class FakeFont:
    def __init__(self) -> None:
        self.Size = 18.0


class FakeTextRange:
    def __init__(self, text: str) -> None:
        self.Text = text
        self.Font = FakeFont()

    def Replace(self, find: str, replacement: str) -> "FakeTextRange":
        self.Text = self.Text.replace(find, replacement)
        return self


class FakeTextFrame:
    def __init__(self, text: str) -> None:
        self.HasText = bool(text)
        self.TextRange = FakeTextRange(text)


class FakeShape:
    def __init__(
        self,
        shape_id: int,
        name: str,
        text: str,
        *,
        index: int = 1,
        left: float = 0.0,
        top: float = 0.0,
        width: float = 100.0,
        height: float = 50.0,
    ) -> None:
        self.Id = shape_id
        self.Index = index
        self.Name = name
        self.Type = 14
        self.Left = left
        self.Top = top
        self.Width = width
        self.Height = height
        self.HasTextFrame = True
        self.TextFrame = FakeTextFrame(text)
        self.HasTable = False


class FakeShapes:
    def __init__(self, shapes: list[FakeShape], title: FakeShape | None) -> None:
        self._shapes = shapes
        self.Count = len(shapes)
        self.Title = title
        self.added: list[tuple] = []

    def __call__(self, ref) -> FakeShape:
        if isinstance(ref, str):
            for shape in self._shapes:
                if shape.Name == ref:
                    return shape
            raise KeyError(ref)
        return self._shapes[ref - 1]

    def AddTextbox(self, orientation: int, left: float, top: float, width: float, height: float) -> FakeShape:
        self.added.append((orientation, left, top, width, height))
        shape = FakeShape(
            900 + len(self._shapes),
            f"TextBox {len(self._shapes) + 1}",
            "",
            index=len(self._shapes) + 1,
            left=left,
            top=top,
            width=width,
            height=height,
        )
        self._shapes.append(shape)
        self.Count = len(self._shapes)
        return shape


class FakeSlide:
    def __init__(self, index: int, title: str, body: str, notes: str = "") -> None:
        title_shape = FakeShape(1, "Title", title, index=1)
        body_shape = FakeShape(2, "Body", body, index=2)
        self.SlideIndex = index
        self.SlideID = 200 + index
        self.CustomLayout = SimpleNamespace(Name="Title and Content")
        self.Shapes = FakeShapes([title_shape, body_shape], title_shape)
        placeholder = SimpleNamespace(TextFrame=FakeTextFrame(notes))
        placeholders = _Caller(lambda n: placeholder)
        self.NotesPage = SimpleNamespace(
            Shapes=SimpleNamespace(Placeholders=placeholders)
        )


class FakeSlides:
    def __init__(self, slides: list[FakeSlide]) -> None:
        self._slides = slides
        self.Count = len(slides)

    def __call__(self, index: int) -> FakeSlide:
        return self._slides[index - 1]

    def Add(self, index: int, layout: int) -> FakeSlide:
        slide = FakeSlide(index, "", "")
        self._slides.insert(index - 1, slide)
        self.Count = len(self._slides)
        return slide


class FakePresentation:
    def __init__(self, slides: list[FakeSlide]) -> None:
        self.Slides = FakeSlides(slides)
        self.PageSetup = SimpleNamespace(SlideWidth=960.0, SlideHeight=540.0)


def make_target(slides: list[FakeSlide], window=True) -> SimpleNamespace:
    presentation = FakePresentation(slides)
    return SimpleNamespace(
        app=AppKind.POWERPOINT,
        document=presentation,
        window=SimpleNamespace(View=SimpleNamespace(Slide=slides[0])) if window else None,
        info={
            "application": "PowerPoint",
            "document": "Deck.pptx",
            "full_name": r"C:\docs\Deck.pptx",
            "slide_index": 1,
        },
    )


def test_slide_returns_title_and_shape_texts() -> None:
    slides = [FakeSlide(1, "Quarterly Review", "Revenue up 10 percent")]
    result = powerpoint.slide(make_target(slides), max_text_chars=100)
    assert result["title"] == "Quarterly Review"
    assert result["slide_index"] == 1
    assert result["slide_count"] == 1
    assert result["shapes"][0]["text"] == "Quarterly Review"
    assert result["shapes"][1]["text"] == "Revenue up 10 percent"


def test_slide_truncates_text() -> None:
    slides = [FakeSlide(1, "abcdefghij", "klmnopqrst")]
    result = powerpoint.slide(make_target(slides), max_text_chars=4)
    assert result["title"] == "abcd"
    assert result["title_truncated"] is True
    assert result["truncated"] is True


def test_slide_requires_a_document_window() -> None:
    slides = [FakeSlide(1, "Title", "Body")]
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.slide(make_target(slides, window=False), max_text_chars=100)
    assert excinfo.value.code == ErrorCode.NO_ACTIVE_SLIDE


def test_document_lists_slide_titles() -> None:
    slides = [FakeSlide(1, "One", ""), FakeSlide(2, "Two", "")]
    result = powerpoint.document(make_target(slides), max_text_chars=100)
    assert result["slide_count"] == 2
    assert result["slide_titles"] == [
        {"slide_index": 1, "title": "One", "truncated": False},
        {"slide_index": 2, "title": "Two", "truncated": False},
    ]


def test_slides_reads_window_with_geometry_and_notes() -> None:
    slides = [
        FakeSlide(1, "One", "body one", notes="n1"),
        FakeSlide(2, "Two", "body two", notes="n2"),
    ]
    result = powerpoint.slides(make_target(slides), start=2, count=5, max_text_chars=200)
    assert result["unit"] == "pt"
    assert result["slide_count"] == 2
    assert result["slide_width"] == 960.0
    assert result["count"] == 1
    entry = result["slides"][0]
    assert entry["slide_index"] == 2
    assert entry["notes"] == "n2"
    shape = entry["shapes"][0]
    assert shape["name"] == "Title"
    assert shape["geometry"]["unit"] == "pt"
    assert shape["geometry"]["width"] == 100.0


def test_slides_truncates_text_and_clamps_count() -> None:
    slides = [FakeSlide(index, f"T{index}", "x" * 30) for index in range(1, 6)]
    result = powerpoint.slides(make_target(slides), start=1, count=1000, max_text_chars=40)
    assert result["truncated"] is True
    assert result["count"] == 5


def test_slides_rejects_bad_start() -> None:
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.slides(
            make_target([FakeSlide(1, "T", "B")]), start=0, count=5, max_text_chars=100
        )
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_text_replace_dry_run_counts_occurrences() -> None:
    target = make_target([FakeSlide(1, "Title", "Wang 1979 and Wang 1979")])
    result = powerpoint.text_replace(
        target,
        slide_index=1,
        shape_ref=2,
        find="Wang",
        replacement="Williams",
        dry_run=True,
        max_text_chars=100,
        params={},
    )
    assert result["dry_run"] is True
    assert result["occurrences"] == 2
    assert result["changed"] is False
    assert target.document.Slides(1).Shapes(2).TextFrame.TextRange.Text == (
        "Wang 1979 and Wang 1979"
    )


def test_text_replace_applies_and_verifies() -> None:
    target = make_target([FakeSlide(1, "Title", "Sutherland 1963")])
    result = powerpoint.text_replace(
        target,
        slide_index=1,
        shape_ref="Body",
        find="1963",
        replacement="1968",
        dry_run=False,
        max_text_chars=100,
        params={"confirm": True},
    )
    assert result["changed"] is True
    assert result["verified"] is True
    assert "1968" in result["text_preview"]
    assert target.document.Slides(1).Shapes(2).TextFrame.TextRange.Text == (
        "Sutherland 1968"
    )


def test_text_replace_requires_confirmation() -> None:
    target = make_target([FakeSlide(1, "Title", "abc")])
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.text_replace(
            target,
            slide_index=1,
            shape_ref=2,
            find="a",
            replacement="b",
            dry_run=False,
            max_text_chars=100,
            params={},
        )
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED


def test_text_replace_reports_missing_text() -> None:
    target = make_target([FakeSlide(1, "Title", "abc")])
    result = powerpoint.text_replace(
        target,
        slide_index=1,
        shape_ref=2,
        find="zzz",
        replacement="y",
        dry_run=False,
        max_text_chars=100,
        params={"confirm": True},
    )
    assert result["changed"] is False
    assert result["verified"] is False
    assert "note" in result


def test_text_replace_unknown_slide_is_com_error() -> None:
    target = make_target([FakeSlide(1, "Title", "Body")])
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.text_replace(
            target,
            slide_index=9,
            shape_ref=2,
            find="a",
            replacement="b",
            dry_run=True,
            max_text_chars=100,
            params={},
        )
    assert excinfo.value.code == ErrorCode.COM_ERROR


def test_textbox_add_dry_run_checks_fit() -> None:
    target = make_target([FakeSlide(1, "T", "B")])
    result = powerpoint.textbox_add(
        target,
        slide_index=1,
        left=900,
        top=500,
        width=100,
        height=100,
        text="note",
        font_size=None,
        dry_run=True,
        max_text_chars=100,
        params={},
    )
    assert result["fits_in_slide"] is False
    assert "warning" in result
    assert target.document.Slides(1).Shapes.added == []


def test_textbox_add_returns_created_shape() -> None:
    target = make_target([FakeSlide(1, "T", "B")])
    result = powerpoint.textbox_add(
        target,
        slide_index=1,
        left=10,
        top=20,
        width=300,
        height=80,
        text="top note",
        font_size=14,
        dry_run=False,
        max_text_chars=100,
        params={"confirm": True},
    )
    assert result["added"] is True
    assert result["verified"] is True
    shape = result["shape"]
    assert shape["index"] == 3
    assert shape["name"].startswith("TextBox")
    assert shape["geometry"]["unit"] == "pt"
    assert target.document.Slides(1).Shapes.added == [(1, 10.0, 20.0, 300.0, 80.0)]
    created = target.document.Slides(1).Shapes(3)
    assert created.TextFrame.TextRange.Text == "top note"
    assert created.TextFrame.TextRange.Font.Size == 14.0


def test_textbox_add_rejects_text_over_limit() -> None:
    target = make_target([FakeSlide(1, "T", "B")])
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.textbox_add(
            target,
            slide_index=1,
            left=0,
            top=0,
            width=10,
            height=10,
            text="x" * 20,
            font_size=None,
            dry_run=True,
            max_text_chars=10,
            params={},
        )
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_slide_add_appends_and_verifies() -> None:
    target = make_target([FakeSlide(1, "One", ""), FakeSlide(2, "Two", "")])
    result = powerpoint.slide_add(target, index=None, layout=None, dry_run=False, params={"confirm": True})
    assert result["added"] is True
    assert result["slide_index"] == 3
    assert result["layout"] == 2
    assert result["verified"] is True
    assert result["slide_count_after"] == 3


def test_slide_add_rejects_index_beyond_end() -> None:
    target = make_target([FakeSlide(1, "One", "")])
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.slide_add(target, index=5, layout=None, dry_run=True, params={})
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_slide_add_requires_confirmation() -> None:
    target = make_target([FakeSlide(1, "One", "")])
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.slide_add(target, index=None, layout=None, dry_run=False, params={})
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED
