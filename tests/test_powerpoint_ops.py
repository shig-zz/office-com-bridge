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
        self.Name = "Calibri"
        self.Size = 18.0
        self.Bold = 0
        self.Color = SimpleNamespace(RGB=0)


class FakeParagraphFont:
    """Font of a paragraph slice: writes fan out, reads need a uniform value."""

    def __init__(self, text_range: "FakeTextRange", start: int, count: int) -> None:
        self._range = text_range
        self._start = start
        self._count = count

    def _paragraphs(self) -> list[FakeFont]:
        return [self._range._font_for(index) for index in range(self._start, self._start + self._count)]

    def _uniform(self, attribute: str) -> Any:
        values = [getattr(font, attribute) for font in self._paragraphs()]
        unique = set(values)
        return values[0] if len(unique) == 1 else None

    @property
    def Size(self) -> Any:
        return self._uniform("Size")

    @Size.setter
    def Size(self, value: float) -> None:
        for font in self._paragraphs():
            font.Size = value

    @property
    def Name(self) -> Any:
        return self._uniform("Name")

    @Name.setter
    def Name(self, value: str) -> None:
        for font in self._paragraphs():
            font.Name = value

    @property
    def Bold(self) -> Any:
        return self._uniform("Bold")

    @Bold.setter
    def Bold(self, value: int) -> None:
        for font in self._paragraphs():
            font.Bold = value

    @property
    def Color(self) -> Any:
        return SimpleNamespace(RGB=self._uniform_rgb())

    def _uniform_rgb(self) -> Any:
        values = [font.Color.RGB for font in self._paragraphs()]
        unique = set(values)
        return values[0] if len(unique) == 1 else None

    @Color.setter
    def Color(self, _value: int) -> None:  # pragma: no cover - RGB is written instead
        pass


class FakeParagraphRange:
    """TextRange.Paragraphs(start[, count]) result: level and font per paragraph."""

    def __init__(self, text_range: "FakeTextRange", start: int, count: int) -> None:
        self._range = text_range
        self._start = start
        self._count = count

    @property
    def Text(self) -> str:
        return "\r".join(self._range._paragraph_texts()[self._start - 1 : self._start - 1 + self._count])

    @property
    def IndentLevel(self) -> Any:
        values = [self._range._level_for(index) for index in range(self._start, self._start + self._count)]
        unique = set(values)
        return values[0] if len(unique) == 1 else None

    @IndentLevel.setter
    def IndentLevel(self, value: int) -> None:
        for index in range(self._start, self._start + self._count):
            self._range.levels[index] = int(value)

    @property
    def Font(self) -> FakeParagraphFont:
        return FakeParagraphFont(self._range, self._start, self._count)


class FakeTextRange:
    def __init__(self, text: str) -> None:
        self.Text = text
        self.Font = FakeFont()
        self.levels: dict[int, int] = {}
        self.paragraph_fonts: dict[int, FakeFont] = {}

    def _paragraph_texts(self) -> list[str]:
        return str(self.Text).split("\r")

    def _level_for(self, index: int) -> int:
        return self.levels.get(index, 1)

    def _font_for(self, index: int) -> FakeFont:
        if index not in self.paragraph_fonts:
            shared = FakeFont()
            shared.Name = self.Font.Name
            shared.Size = self.Font.Size
            shared.Bold = self.Font.Bold
            shared.Color = SimpleNamespace(RGB=self.Font.Color.RGB)
            self.paragraph_fonts[index] = shared
        return self.paragraph_fonts[index]

    def Paragraphs(self, start: int | None = None, count: int | None = None):
        if start is None:
            return SimpleNamespace(Count=len(self._paragraph_texts()))
        return FakeParagraphRange(self, start, count or 1)

    def Replace(self, find: str, replacement: str) -> "FakeTextRange":
        self.Text = self.Text.replace(find, replacement)
        return self


class FakeTextFrame:
    def __init__(self, text: str) -> None:
        self.HasText = bool(text)
        self.TextRange = FakeTextRange(text)
        self.AutoSize = 0
        self.WordWrap = -1


class FakeTextFrame2:
    """TextFrame2 whose AutoSize=1 (shape fits text) re-measures the shape."""

    def __init__(self, shape: "FakeShape", text_range: FakeTextRange) -> None:
        self._shape = shape
        self._autosize = 0
        self.TextRange = text_range
        self.WordWrap = -1

    @property
    def AutoSize(self) -> int:
        return self._autosize

    @AutoSize.setter
    def AutoSize(self, value: int) -> None:
        self._autosize = int(value)
        if self._autosize == 1:
            self._shape.grow_to_fit_text()


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
        has_text_frame: bool = True,
    ) -> None:
        self.Id = shape_id
        self.Index = index
        self.Name = name
        self.Type = 14
        self.Left = left
        self.Top = top
        self.Width = width
        self.Height = height
        self.HasTextFrame = has_text_frame
        self.TextFrame = FakeTextFrame(text if has_text_frame else "")
        self.TextFrame2 = FakeTextFrame2(self, self.TextFrame.TextRange)
        self.PlaceholderFormat = SimpleNamespace(Type=2)
        self.HasTable = False
        self._shapes_container = None

    def grow_to_fit_text(self) -> None:
        """Approximate PowerPoint's 'resize shape to fit text' for measurement tests."""
        text = str(self.TextFrame.TextRange.Text)
        size = float(self.TextFrame.TextRange.Font.Size or 12)
        per_line = max(1, int(self.Width / (size * 0.6)))
        lines = 0
        for paragraph in text.split("\r"):
            lines += max(1, -(-len(paragraph) // per_line))
        self.Height = round(lines * size * 1.2, 2)

    def Duplicate(self) -> FakeShape:
        if self._shapes_container is None:
            raise AssertionError("this shape is not attached to a slide")
        return self._shapes_container.duplicate_shape(self)

    def Delete(self) -> None:
        if self._shapes_container is not None:
            self._shapes_container.remove_shape(self)


class FakeShapes:
    def __init__(self, shapes: list[FakeShape], title: FakeShape | None) -> None:
        self._shapes = shapes
        self.Count = len(shapes)
        self.Title = title
        self.added: list[tuple] = []
        for shape in shapes:
            shape._shapes_container = self  # noqa: SLF001

    def append(self, shape: FakeShape) -> None:
        self._shapes.append(shape)
        self.Count = len(self._shapes)
        shape._shapes_container = self  # noqa: SLF001

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
        shape._shapes_container = self  # noqa: SLF001
        return shape

    def duplicate_shape(self, shape: FakeShape) -> FakeShape:
        copy = FakeShape(
            800 + len(self._shapes),
            f"{shape.Name} 2",
            str(shape.TextFrame.TextRange.Text),
            index=len(self._shapes) + 1,
            left=shape.Left,
            top=shape.Top,
            width=shape.Width,
            height=shape.Height,
            has_text_frame=shape.HasTextFrame,
        )
        copy.TextFrame.TextRange.Font.Size = shape.TextFrame.TextRange.Font.Size
        copy.TextFrame.TextRange.Font.Name = shape.TextFrame.TextRange.Font.Name
        self._shapes.append(copy)
        self.Count = len(self._shapes)
        copy._shapes_container = self  # noqa: SLF001
        return copy

    def remove_shape(self, shape: FakeShape) -> None:
        self._shapes.remove(shape)
        self.Count = len(self._shapes)


class FakeSlide:
    def __init__(self, index: int, title: str, body: str, notes: str = "") -> None:
        title_shape = FakeShape(1, "Title", title, index=1)
        body_shape = FakeShape(2, "Body", body, index=2)
        self.SlideIndex = index
        self.SlideID = 200 + index
        self.CustomLayout = SimpleNamespace(Name="Title and Content")
        self.Shapes = FakeShapes([title_shape, body_shape], title_shape)
        self.notes_range = FakeTextRange(notes)
        self.NotesPage = SimpleNamespace(
            Shapes=SimpleNamespace(
                Placeholders=_Caller(lambda n: SimpleNamespace(TextFrame=SimpleNamespace(TextRange=self.notes_range)))
            )
        )
        self._container = None

    def Delete(self) -> None:
        if self._container is not None:
            self._container.remove_slide(self)

    def MoveTo(self, to: int) -> None:
        if self._container is not None:
            self._container.move_slide(self, to)


class FakeSlides:
    def __init__(self, slides: list[FakeSlide]) -> None:
        self._slides = slides
        self.Count = len(slides)
        for slide in slides:
            slide._container = self  # noqa: SLF001

    def __call__(self, index: int) -> FakeSlide:
        return self._slides[index - 1]

    def Add(self, index: int, layout: int) -> FakeSlide:
        slide = FakeSlide(index, "", "")
        slide._container = self  # noqa: SLF001
        self._slides.insert(index - 1, slide)
        self.Count = len(self._slides)
        return slide

    def remove_slide(self, slide: FakeSlide) -> None:
        self._slides.remove(slide)
        self.Count = len(self._slides)

    def move_slide(self, slide: FakeSlide, to: int) -> None:
        self._slides.remove(slide)
        self._slides.insert(to - 1, slide)
        for position, item in enumerate(self._slides, start=1):
            item.SlideIndex = position


class FakeSectionProperties:
    """PowerPoint's parameterized SectionProperties object (Name/FirstSlide/... take an index)."""

    def __init__(self, sections: list[tuple[str, int, int]]) -> None:
        self._sections = sections
        self.Count = len(sections)

    def Name(self, index: int) -> str:
        return self._sections[index - 1][0]

    def FirstSlide(self, index: int) -> int:
        return self._sections[index - 1][1]

    def SlidesCount(self, index: int) -> int:
        return self._sections[index - 1][2]

    def SectionID(self, index: int) -> str:
        return f"{{SEC{index:04d}}}"


class FakePresentation:
    def __init__(self, slides: list[FakeSlide], sections: list[tuple[str, int, int]] | None = None) -> None:
        self.Slides = FakeSlides(slides)
        self.PageSetup = SimpleNamespace(SlideWidth=960.0, SlideHeight=540.0)
        if sections is not None:
            self.SectionProperties = FakeSectionProperties(sections)


def make_target(
    slides: list[FakeSlide],
    window=True,
    sections: list[tuple[str, int, int]] | None = None,
) -> SimpleNamespace:
    presentation = FakePresentation(slides, sections=sections)
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


def test_notes_reads_speaker_notes() -> None:
    target = make_target([FakeSlide(1, "One", "body", notes="note text")])
    result = powerpoint.notes(target, slide_index=1, max_text_chars=100)
    assert result["slide"] == 1
    assert result["notes"] == "note text"
    assert result["characters"] == 9
    assert result["truncated"] is False


def test_notes_set_dry_run_previews_current_text() -> None:
    target = make_target([FakeSlide(1, "One", "body", notes="old")])
    result = powerpoint.notes_set(
        target,
        slide_index=1,
        text="new",
        dry_run=True,
        max_text_chars=100,
        params={},
    )
    assert result["notes_preview"] == "old"
    assert result["changed"] is False
    assert target.document.Slides(1).notes_range.Text == "old"


def test_notes_set_applies_and_verifies() -> None:
    target = make_target([FakeSlide(1, "One", "body", notes="old")])
    result = powerpoint.notes_set(
        target,
        slide_index=1,
        text="first line\nsecond line",
        dry_run=False,
        max_text_chars=100,
        params={"confirm": True},
    )
    assert result["changed"] is True
    assert result["verified"] is True
    # PowerPoint separates paragraphs with CR, so the LF is converted rather than
    # written as an ordinary character inside one paragraph.
    assert target.document.Slides(1).notes_range.Text == "first line\rsecond line"
    assert result["paragraph_breaks"] == 1
    assert result["newline_conversions"] == 1


def test_notes_set_requires_confirmation() -> None:
    target = make_target([FakeSlide(1, "One", "body", notes="old")])
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.notes_set(
            target,
            slide_index=1,
            text="new",
            dry_run=False,
            max_text_chars=100,
            params={},
        )
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED


def test_slide_delete_dry_run_reports_title() -> None:
    target = make_target([FakeSlide(1, "Keep", ""), FakeSlide(2, "Drop", "")])
    result = powerpoint.slide_delete(target, slide_index=2, dry_run=True, params={})
    assert result["dry_run"] is True
    assert result["slide_title"] == "Drop"
    assert result["deleted"] is False
    assert target.document.Slides.Count == 2


def test_slide_delete_removes_and_verifies() -> None:
    target = make_target([FakeSlide(1, "Keep", ""), FakeSlide(2, "Drop", "")])
    result = powerpoint.slide_delete(target, slide_index=2, dry_run=False, params={"confirm": True})
    assert result["deleted"] is True
    assert result["verified"] is True
    assert result["slide_count_after"] == 1
    assert target.document.Slides(1).SlideIndex == 1


def test_slide_delete_requires_confirmation() -> None:
    target = make_target([FakeSlide(1, "One", "")])
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.slide_delete(target, slide_index=1, dry_run=False, params={})
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED


def test_slide_delete_result_carries_undo_hint() -> None:
    target = make_target([FakeSlide(1, "Keep", ""), FakeSlide(2, "Drop", "")])
    result = powerpoint.slide_delete(target, slide_index=2, dry_run=False, params={"confirm": True})
    assert "undo" in result["undo_hint"]


# ---------------------------------------------------------------------------
# Phase 3C: slide move, shape delete, shape set-format


def test_slide_move_dry_run_reports_plan() -> None:
    target = make_target(
        [FakeSlide(1, "First", ""), FakeSlide(2, "Second", ""), FakeSlide(3, "Note", "")]
    )
    result = powerpoint.slide_move(target, slide_index=3, to_index=1, dry_run=True, params={})
    assert result["direction"] == "up"
    assert result["slide_title"] == "Note"
    assert result["moved"] is False
    assert target.document.Slides(1).SlideID == 201


def test_slide_move_down_applies_and_verifies() -> None:
    target = make_target(
        [FakeSlide(1, "First", ""), FakeSlide(2, "Second", ""), FakeSlide(3, "Note", "")]
    )
    result = powerpoint.slide_move(target, slide_index=1, to_index=3, dry_run=False, params={"confirm": True})
    assert result["moved"] is True
    assert result["verified"] is True
    moved = target.document.Slides(3)
    assert moved.SlideID == 201
    assert moved.SlideIndex == 3
    assert result["slide_title_after"] == "First"


def test_slide_move_up_applies_and_verifies() -> None:
    target = make_target(
        [FakeSlide(1, "First", ""), FakeSlide(2, "Second", ""), FakeSlide(3, "Note", "")]
    )
    result = powerpoint.slide_move(target, slide_index=3, to_index=1, dry_run=False, params={"confirm": True})
    assert result["verified"] is True
    assert target.document.Slides(1).SlideID == 203
    assert target.document.Slides(2).SlideID == 201


def test_slide_move_requires_confirmation() -> None:
    target = make_target([FakeSlide(1, "One", ""), FakeSlide(2, "Two", "")])
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.slide_move(target, slide_index=1, to_index=2, dry_run=False, params={})
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED
    assert target.document.Slides(1).SlideID == 201


def test_slide_move_rejects_bad_targets() -> None:
    target = make_target([FakeSlide(1, "One", ""), FakeSlide(2, "Two", "")])
    for bad_to in (0, 3, 1, "2"):
        with pytest.raises(BridgeError) as excinfo:
            powerpoint.slide_move(target, slide_index=1, to_index=bad_to, dry_run=True, params={})
        assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_shape_delete_dry_run_reports_preview() -> None:
    target = make_target([FakeSlide(1, "Title", "drop me")])
    result = powerpoint.shape_delete(
        target, slide_index=1, shape_ref=2, dry_run=True, max_text_chars=100, params={}
    )
    assert result["shape_name"] == "Body"
    assert result["text_preview"] == "drop me"
    assert result["shape_count_before"] == 2
    assert result["deleted"] is False
    assert target.document.Slides(1).Shapes.Count == 2


def test_shape_delete_removes_and_verifies() -> None:
    target = make_target([FakeSlide(1, "Title", "drop me")])
    result = powerpoint.shape_delete(
        target, slide_index=1, shape_ref="Body", dry_run=False, max_text_chars=100, params={"confirm": True}
    )
    assert result["deleted"] is True
    assert result["verified"] is True
    assert result["shape_count_after"] == 1
    assert "undo" in result["undo_hint"]


def test_shape_delete_requires_confirmation() -> None:
    target = make_target([FakeSlide(1, "Title", "body")])
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.shape_delete(
            target, slide_index=1, shape_ref=2, dry_run=False, max_text_chars=100, params={}
        )
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED
    assert target.document.Slides(1).Shapes.Count == 2


def test_shape_set_format_dry_run_previews_current_font() -> None:
    target = make_target([FakeSlide(1, "Title", "body text")])
    result = powerpoint.shape_set_format(
        target,
        slide_index=1,
        shape_ref=2,
        font_size=14,
        font_name=None,
        color="FF8000",
        bold=True,
        dry_run=True,
        params={},
    )
    assert result["apply_to"] == "text_range"
    assert result["font_before"] == {
        "name": "Calibri",
        "size": 18.0,
        "bold": False,
        "color": "000000",
    }
    assert result["planned"] == {"font_size": 14.0, "color": "FF8000", "bold": True}
    assert result["changed"] is False
    assert target.document.Slides(1).Shapes(2).TextFrame.TextRange.Font.Size == 18.0


def test_shape_set_format_applies_and_verifies() -> None:
    target = make_target([FakeSlide(1, "Title", "body text")])
    result = powerpoint.shape_set_format(
        target,
        slide_index=1,
        shape_ref=2,
        font_size=14,
        font_name="Meiryo",
        color="FF8000",
        bold=False,
        dry_run=False,
        params={"confirm": True},
    )
    assert result["changed"] is True
    assert result["verified"] is True
    font = target.document.Slides(1).Shapes(2).TextFrame.TextRange.Font
    assert font.Size == 14.0
    assert font.Name == "Meiryo"
    assert font.Bold == 0
    # FF8000 is stored as the BGR integer Office uses
    assert font.Color.RGB == 0x0080FF
    assert result["font_after"]["color"] == "FF8000"


def test_shape_set_format_requires_confirmation() -> None:
    target = make_target([FakeSlide(1, "Title", "body text")])
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.shape_set_format(
            target, slide_index=1, shape_ref=2, font_size=12, font_name=None, color=None, bold=None,
            dry_run=False, params={},
        )
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED


def test_shape_set_format_rejects_bad_parameters() -> None:
    target = make_target([FakeSlide(1, "Title", "body text")])
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.shape_set_format(
            target, slide_index=1, shape_ref=2, font_size=None, font_name=None, color=None, bold=None,
            dry_run=True, params={},
        )
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST
    for bad_color in ("red", "GGGGGG", 5):
        with pytest.raises(BridgeError) as excinfo:
            powerpoint.shape_set_format(
                target, slide_index=1, shape_ref=2, font_size=None, font_name=None,
                color=bad_color, bold=None, dry_run=True, params={},
            )
        assert excinfo.value.code == ErrorCode.INVALID_REQUEST
    for bad_size in (0, 500, "12"):
        with pytest.raises(BridgeError) as excinfo:
            powerpoint.shape_set_format(
                target, slide_index=1, shape_ref=2, font_size=bad_size, font_name=None,
                color=None, bold=None, dry_run=True, params={},
            )
        assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_shape_set_format_refuses_shape_without_text() -> None:
    target = make_target([FakeSlide(1, "Title", "")])
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.shape_set_format(
            target, slide_index=1, shape_ref=2, font_size=12, font_name=None, color=None, bold=None,
            dry_run=True, params={},
        )
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_shape_set_format_reports_missing_text_frame() -> None:
    target = make_target([FakeSlide(1, "Title", "body")])
    target.document.Slides(1).Shapes.append(FakeShape(3, "Picture 3", "", has_text_frame=False))
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.shape_set_format(
            target, slide_index=1, shape_ref=3, font_size=12, font_name=None, color=None, bold=None,
            dry_run=True, params={},
        )
    assert excinfo.value.code == ErrorCode.TEXT_FRAME_MISSING
    assert excinfo.value.details["has_text_frame"] == 0
    assert excinfo.value.details["shape_name"] == "Picture 3"


# ---------------------------------------------------------------------------
# text read: one call instead of a chain of com gets
# ---------------------------------------------------------------------------


def test_text_read_reports_paragraphs_geometry_and_autofit() -> None:
    target = make_target([FakeSlide(1, "Title", "alpha\rbeta")])
    shape = target.document.Slides(1).Shapes(2)
    shape.Top = 75.7
    shape.Height = 442.25
    shape.TextFrame2.AutoSize = 2
    result = powerpoint.text_read(target, slide_index=1, shape_ref=2, max_text_chars=100)
    assert result["shape_name"] == "Body"
    assert result["placeholder_type"] == 2
    assert result["characters"] == 10
    assert result["paragraph_count"] == 2
    assert [entry["paragraph"] for entry in result["paragraphs"]] == [1, 2]
    assert result["paragraphs"][1]["font_size"] == 18.0
    assert result["paragraphs"][1]["indent_level"] == 1
    assert result["geometry"]["height"] == 442.25
    assert result["autofit"]["autofit"] == "text_fits_shape"
    assert result["autofit"]["text_frame2_autosize"] == 2


def test_text_read_reports_missing_text_frame() -> None:
    target = make_target([FakeSlide(1, "Title", "body")])
    target.document.Slides(1).Shapes.append(FakeShape(3, "Picture 3", "", has_text_frame=False))
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.text_read(target, slide_index=1, shape_ref=3, max_text_chars=100)
    assert excinfo.value.code == ErrorCode.TEXT_FRAME_MISSING
    assert excinfo.value.details["shape_type"] == 14


# ---------------------------------------------------------------------------
# text set: whole-frame rewrite with levels and sizes
# ---------------------------------------------------------------------------


def test_text_set_dry_run_previews_text_and_levels() -> None:
    target = make_target([FakeSlide(1, "Title", "old\rtext")])
    result = powerpoint.text_set(
        target,
        slide_index=1,
        shape_ref=2,
        paragraphs=[{"text": "０．表紙"}, {"text": "サブ項目", "level": 2, "font_size": 13}],
        dry_run=True,
        max_text_chars=200,
        params={},
    )
    assert result["dry_run"] is True
    assert result["text_preview"] == "old\rtext"
    assert result["planned_text"] == "０．表紙\rサブ項目"
    assert result["paragraphs_before"] == 2
    assert result["paragraphs_after"] == 2
    assert [item.get("indent_level") for item in result["planned"]] == [None, 2]
    assert target.document.Slides(1).Shapes(2).TextFrame.TextRange.Text == "old\rtext"


def test_text_set_applies_levels_sizes_and_normalizes_line_feeds() -> None:
    target = make_target([FakeSlide(1, "Title", "old")])
    result = powerpoint.text_set(
        target,
        slide_index=1,
        shape_ref=2,
        paragraphs=[
            "first",
            {"text": "second", "level": 2, "font_size": 13.0},
            {"text": "third\nmore", "font_size": 16},
        ],
        dry_run=False,
        max_text_chars=200,
        params={"confirm": True},
    )
    text_range = target.document.Slides(1).Shapes(2).TextFrame.TextRange
    assert text_range.Text == "first\rsecond\rthird\rmore"
    assert result["verified"] is True
    assert result["paragraphs_after"] == 4
    assert result["newline_conversions"] == 1
    assert text_range.Paragraphs(2).IndentLevel == 2
    assert text_range.Paragraphs(2).Font.Size == 13.0
    assert text_range.Paragraphs(3, 2).Font.Size == 16.0
    assert text_range.Paragraphs(1).IndentLevel == 1


def test_text_set_writes_into_empty_placeholder() -> None:
    target = make_target([FakeSlide(1, "Title", "")])
    result = powerpoint.text_set(
        target,
        slide_index=1,
        shape_ref=2,
        paragraphs=["０．表紙", "１．本体"],
        dry_run=False,
        max_text_chars=200,
        params={"confirm": True},
    )
    assert result["verified"] is True
    assert result["characters_before"] == 0
    assert target.document.Slides(1).Shapes(2).TextFrame.TextRange.Text == "０．表紙\r１．本体"


def test_text_set_requires_confirmation() -> None:
    target = make_target([FakeSlide(1, "Title", "old")])
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.text_set(
            target,
            slide_index=1,
            shape_ref=2,
            paragraphs=["new"],
            dry_run=False,
            max_text_chars=200,
            params={},
        )
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED
    assert target.document.Slides(1).Shapes(2).TextFrame.TextRange.Text == "old"


def test_text_set_rejects_bad_paragraph_entries() -> None:
    target = make_target([FakeSlide(1, "Title", "old")])
    for bad in ([], [5], [{"text": 5}], [{"text": "x", "oops": 1}], [{"text": "x", "level": 0}]):
        with pytest.raises(BridgeError) as excinfo:
            powerpoint.text_set(
                target,
                slide_index=1,
                shape_ref=2,
                paragraphs=bad,
                dry_run=True,
                max_text_chars=200,
                params={},
            )
        assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_text_set_rejects_shape_without_text_frame() -> None:
    target = make_target([FakeSlide(1, "Title", "body")])
    target.document.Slides(1).Shapes.append(FakeShape(3, "Picture 3", "", has_text_frame=False))
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.text_set(
            target,
            slide_index=1,
            shape_ref=3,
            paragraphs=["x"],
            dry_run=False,
            max_text_chars=200,
            params={"confirm": True},
        )
    assert excinfo.value.code == ErrorCode.TEXT_FRAME_MISSING


# ---------------------------------------------------------------------------
# text set_format: paragraph ranges and indent levels
# ---------------------------------------------------------------------------


def test_text_set_format_applies_to_paragraph_range() -> None:
    target = make_target([FakeSlide(1, "Title", "a\rb\rc\rd")])
    result = powerpoint.text_set_format(
        target,
        slide_index=1,
        shape_ref=2,
        paragraph="2-3",
        font_size=13,
        font_name=None,
        color=None,
        bold=None,
        indent_level=2,
        dry_run=False,
        params={"confirm": True},
    )
    text_range = target.document.Slides(1).Shapes(2).TextFrame.TextRange
    assert result["apply_to"] == "paragraphs"
    assert result["paragraph"] == "2-3"
    assert result["verified"] is True
    assert text_range.Paragraphs(2, 2).IndentLevel == 2
    assert text_range.Paragraphs(2, 2).Font.Size == 13.0
    assert text_range.Paragraphs(1).IndentLevel == 1
    assert text_range.Paragraphs(1).Font.Size == 18.0
    assert text_range.Paragraphs(4).Font.Size == 18.0


def test_text_set_format_dry_run_reports_each_block() -> None:
    target = make_target([FakeSlide(1, "Title", "a\rb\rc")])
    result = powerpoint.text_set_format(
        target,
        slide_index=1,
        shape_ref=2,
        paragraph=["1", "3"],
        font_size=None,
        font_name=None,
        color=None,
        bold=None,
        indent_level=3,
        dry_run=True,
        params={},
    )
    assert [block["target"] for block in result["blocks"]] == [
        "paragraphs(1)",
        "paragraphs(3)",
    ]
    assert result["planned"] == {"indent_level": 3}
    assert target.document.Slides(1).Shapes(2).TextFrame.TextRange.levels == {}


def test_text_set_format_validates_parameters() -> None:
    target = make_target([FakeSlide(1, "Title", "a\rb")])
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.text_set_format(
            target, slide_index=1, shape_ref=2, paragraph=None, font_size=None, font_name=None,
            color=None, bold=None, indent_level=None, dry_run=True, params={},
        )
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST
    for bad_paragraph in ("0", "5-2", "abc"):
        with pytest.raises(BridgeError) as excinfo:
            powerpoint.text_set_format(
                target, slide_index=1, shape_ref=2, paragraph=bad_paragraph, font_size=12,
                font_name=None, color=None, bold=None, indent_level=None, dry_run=True, params={},
            )
        assert excinfo.value.code == ErrorCode.INVALID_REQUEST


# ---------------------------------------------------------------------------
# autofit: set and read
# ---------------------------------------------------------------------------


def test_text_set_autofit_text_goes_through_textframe2() -> None:
    target = make_target([FakeSlide(1, "Title", "body")])
    result = powerpoint.text_set_autofit(
        target, slide_index=1, shape_ref=2, behavior="text", dry_run=False, params={"confirm": True}
    )
    shape = target.document.Slides(1).Shapes(2)
    assert shape.TextFrame2.AutoSize == 2
    assert shape.TextFrame.AutoSize == 0
    assert result["applied_via"] == "TextFrame2.AutoSize"
    assert result["verified"] is True
    assert result["autofit_after"]["autofit"] == "text_fits_shape"
    assert result["attempts"][0] == {"member": "TextFrame2.AutoSize", "ok": True, "value": 2}


def test_text_set_autofit_shape_falls_back_to_textframe_and_cast_is_logged() -> None:
    target = make_target([FakeSlide(1, "Title", "body")])
    shape = target.document.Slides(1).Shapes(2)
    del shape.TextFrame2

    class FixedFrame2:
        AutoSize = 1
        WordWrap = -1

    shape.TextFrame2 = None
    del shape.TextFrame2
    result = powerpoint.text_set_autofit(
        target, slide_index=1, shape_ref=2, behavior="shape", dry_run=False, params={"confirm": True}
    )
    assert shape.TextFrame.AutoSize == 1
    assert result["applied_via"] == "TextFrame.AutoSize"
    assert "TextFrame2.AutoSize (type-library cast)" in [entry["member"] for entry in result["attempts"]]
    assert result["verified"] is True


def test_text_set_autofit_refused_when_no_member_accepts_it() -> None:
    class RefusingTextFrame:
        HasText = -1
        WordWrap = -1
        AutoSize = 0

        def __setattr__(self, key: str, value: object) -> None:
            if key == "AutoSize":
                raise RuntimeError("Office refused the autofit value (DISP_E_EXCEPTION)")
            object.__setattr__(self, key, value)

    target = make_target([FakeSlide(1, "Title", "body")])
    shape = target.document.Slides(1).Shapes(2)
    del shape.TextFrame2
    shape.TextFrame = RefusingTextFrame()
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.text_set_autofit(
            target, slide_index=1, shape_ref=2, behavior="text", dry_run=False, params={"confirm": True}
        )
    error = excinfo.value
    assert error.code == ErrorCode.AUTOFIT_REFUSED
    assert error.details["behavior"] == "text"
    assert all(not entry["ok"] for entry in error.details["attempts"])
    assert error.details["autosize_enums"]["TextFrame2.AutoSize (MsoAutoSize)"]["text_fits_shape"] == 2
    assert "hint" in error.details


def test_autofit_read_reports_both_members_and_enums() -> None:
    target = make_target([FakeSlide(1, "Title", "body")])
    shape = target.document.Slides(1).Shapes(2)
    shape.TextFrame.AutoSize = 1
    shape.TextFrame2.AutoSize = 2
    result = powerpoint.autofit_read(target, slide_index=1, shape_ref=2)
    assert result["autofit"]["text_frame_autosize"] == 1
    assert result["autofit"]["text_frame2_autosize"] == 2
    assert result["autofit"]["autofit"] == "text_fits_shape"
    assert result["autofit"]["autofit_member"] == "TextFrame2.AutoSize"
    assert result["autosize_enums"]["TextFrame.AutoSize (PpAutoSize)"]["shape_fits_text"] == 1


def test_text_set_autofit_validates_behavior() -> None:
    target = make_target([FakeSlide(1, "Title", "body")])
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.text_set_autofit(
            target, slide_index=1, shape_ref=2, behavior="shrink", dry_run=True, params={}
        )
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST
    assert excinfo.value.details["values"] == ["none", "shape", "text"]


# ---------------------------------------------------------------------------
# text measure: duplicate, grow, read, delete
# ---------------------------------------------------------------------------


def test_text_measure_reports_required_height_and_cleans_up() -> None:
    target = make_target([FakeSlide(1, "Title", "x" * 400)])
    shape = target.document.Slides(1).Shapes(2)
    shape.TextFrame.TextRange.Font.Size = 20
    result = powerpoint.text_measure(
        target, slide_index=1, shape_ref=2, text=None, font_size=None, dry_run=False,
        params={"confirm": True},
    )
    assert result["measured"] is True
    assert result["current_height"] == 50.0
    assert result["required_height"] > 50.0
    assert result["fits"] is False
    assert result["margin"] < 0
    assert result["duplicate_deleted"] is True
    assert target.document.Slides(1).Shapes.Count == 2
    assert shape.Height == 50.0


def test_text_measure_smaller_font_needs_less_height() -> None:
    target = make_target([FakeSlide(1, "Title", "short body")])
    big = powerpoint.text_measure(
        target, slide_index=1, shape_ref=2, text="長" * 200, font_size=32, dry_run=False,
        params={"confirm": True},
    )
    small = powerpoint.text_measure(
        target, slide_index=1, shape_ref=2, text="長" * 200, font_size=10, dry_run=False,
        params={"confirm": True},
    )
    assert small["required_height"] < big["required_height"]
    assert target.document.Slides(1).Shapes(2).TextFrame.TextRange.Text == "short body"
    assert target.document.Slides(1).Shapes.Count == 2


def test_text_measure_dry_run_does_not_duplicate() -> None:
    target = make_target([FakeSlide(1, "Title", "body")])
    result = powerpoint.text_measure(
        target, slide_index=1, shape_ref=2, text=None, font_size=None, dry_run=True, params={}
    )
    assert result["measured"] is False
    assert result["geometry"]["height"] == 50.0
    assert target.document.Slides(1).Shapes.Count == 2


def test_text_measure_requires_confirmation() -> None:
    target = make_target([FakeSlide(1, "Title", "body")])
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.text_measure(
            target, slide_index=1, shape_ref=2, text=None, font_size=None, dry_run=False, params={}
        )
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED


# ---------------------------------------------------------------------------
# table of contents
# ---------------------------------------------------------------------------


def _toc_slides() -> list[FakeSlide]:
    return [
        FakeSlide(1, "表紙", "subtitle"),
        FakeSlide(2, "改定履歴", "body"),
        FakeSlide(3, "目次", "０．表紙\r１．改定履歴"),
        FakeSlide(4, "部品の種類", "body"),
        FakeSlide(5, "アルミフレーム", "body"),
    ]


def test_toc_fill_dry_run_reuses_fullwidth_numbering() -> None:
    target = make_target(_toc_slides())
    result = powerpoint.toc_fill(
        target,
        slide_index=3,
        shape_ref=2,
        start=None,
        end=None,
        exclude="5",
        levels=None,
        numbering=None,
        font_size=16.0,
        sub_font_size=None,
        dry_run=True,
        max_text_chars=500,
        params={},
    )
    assert result["numbering_style"]["detected"]["digits"] == "fullwidth"
    assert result["numbering_style"]["detected"]["start"] == 0
    assert result["numbering_style"]["separator"] == "．"
    assert result["preview_text"] == "０．表紙\r１．改定履歴\r２．部品の種類"
    assert result["skipped"] == [
        {"slide_index": 3, "reason": "excluded"},
        {"slide_index": 5, "reason": "excluded"},
    ]
    assert target.document.Slides(3).Shapes(2).TextFrame.TextRange.Text == "０．表紙\r１．改定履歴"


def test_toc_fill_applies_levels_and_per_level_sizes() -> None:
    target = make_target(_toc_slides())
    result = powerpoint.toc_fill(
        target,
        slide_index=3,
        shape_ref=2,
        start=None,
        end=None,
        exclude=None,
        levels="4-5=2",
        numbering=None,
        font_size=16.0,
        sub_font_size=13.0,
        dry_run=False,
        max_text_chars=500,
        params={"confirm": True},
    )
    text_range = target.document.Slides(3).Shapes(2).TextFrame.TextRange
    assert text_range.Text == "０．表紙\r１．改定履歴\r部品の種類\rアルミフレーム"
    assert result["entry_count"] == 4
    assert result["verified"] is True
    assert result["skipped"] == [{"slide_index": 3, "reason": "excluded"}]
    assert text_range.Paragraphs(3, 2).IndentLevel == 2
    assert text_range.Paragraphs(3, 2).Font.Size == 13.0
    assert text_range.Paragraphs(1, 2).Font.Size == 16.0
    assert [entry["level"] for entry in result["entries"]] == [1, 1, 2, 2]


def test_toc_fill_auto_levels_children() -> None:
    slides = [
        FakeSlide(1, "目次", "０．"),
        FakeSlide(2, "部品の種類", "アルミフレーム\rパネル"),
        FakeSlide(3, "アルミフレーム", "frame"),
        FakeSlide(4, "パネル", "panel"),
    ]
    target = make_target(slides)
    result = powerpoint.toc_fill(
        target,
        slide_index=1,
        shape_ref=2,
        start=None,
        end=None,
        exclude=None,
        levels="auto",
        numbering=None,
        font_size=None,
        sub_font_size=None,
        dry_run=True,
        max_text_chars=500,
        params={},
    )
    assert result["levels"]["mode"] == "auto"
    assert [(entry["slide_index"], entry["level"]) for entry in result["entries"]] == [
        (2, 1),
        (3, 2),
        (4, 2),
    ]
    assert result["preview_text"] == "０．部品の種類\rアルミフレーム\rパネル"


def test_toc_fill_numbering_none_writes_bare_titles() -> None:
    target = make_target(_toc_slides())
    result = powerpoint.toc_fill(
        target,
        slide_index=3,
        shape_ref=2,
        start=4,
        end=5,
        exclude=None,
        levels=None,
        numbering="none",
        font_size=None,
        sub_font_size=None,
        dry_run=True,
        max_text_chars=500,
        params={},
    )
    assert result["preview_text"] == "部品の種類\rアルミフレーム"
    assert result["numbering_style"]["digits"] == "none"


def test_toc_fill_reports_ascii_numbering_style() -> None:
    slides = _toc_slides()
    slides[2] = FakeSlide(3, "目次", "1. 表紙\r2. 改定履歴")
    target = make_target(slides)
    result = powerpoint.toc_fill(
        target,
        slide_index=3,
        shape_ref=2,
        start=None,
        end=4,
        exclude=None,
        levels=None,
        numbering=None,
        font_size=None,
        sub_font_size=None,
        dry_run=True,
        max_text_chars=500,
        params={},
    )
    assert result["numbering_style"]["digits"] == "ascii"
    assert result["numbering_style"]["start"] == 1
    # The contents slide is never listed in its own table of contents, and the
    # detected separator keeps its trailing space ("1. " rather than "1.").
    assert result["numbering_style"]["separator"] == ". "
    assert result["preview_text"] == "1. 表紙\r2. 改定履歴\r3. 部品の種類"
    assert result["skipped"] == [{"slide_index": 3, "reason": "excluded"}]


def test_toc_fill_requires_confirmation_and_validates_options() -> None:
    target = make_target(_toc_slides())
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.toc_fill(
            target, slide_index=3, shape_ref=2, start=None, end=None, exclude=None, levels=None,
            numbering=None, font_size=None, sub_font_size=None, dry_run=False, max_text_chars=500,
            params={},
        )
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED
    for bad_levels in ("6-12", "abc", "0=2"):
        with pytest.raises(BridgeError) as excinfo:
            powerpoint.toc_fill(
                target, slide_index=3, shape_ref=2, start=None, end=None, exclude=None,
                levels=bad_levels, numbering=None, font_size=None, sub_font_size=None,
                dry_run=True, max_text_chars=500, params={},
            )
        assert excinfo.value.code == ErrorCode.INVALID_REQUEST
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.toc_fill(
            target, slide_index=3, shape_ref=2, start=None, end=None, exclude=None, levels=None,
            numbering="roman", font_size=None, sub_font_size=None, dry_run=True,
            max_text_chars=500, params={},
        )
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_toc_fill_rejects_range_without_titles() -> None:
    target = make_target([FakeSlide(1, "目次", "０．"), FakeSlide(2, "", "no title")])
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.toc_fill(
            target, slide_index=1, shape_ref=2, start=2, end=2, exclude=None, levels=None,
            numbering=None, font_size=None, sub_font_size=None, dry_run=True,
            max_text_chars=500, params={},
        )
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST
    assert excinfo.value.details["skipped"] == [{"slide_index": 2, "reason": "no_title"}]


# ---------------------------------------------------------------------------
# sections
# ---------------------------------------------------------------------------


def test_sections_read_names_and_slide_ranges() -> None:
    slides = [FakeSlide(index, f"T{index}", "body") for index in range(1, 6)]
    target = make_target(slides, sections=[("概要", 1, 2), ("本論", 3, 3)])
    result = powerpoint.sections(target)
    assert result["section_count"] == 2
    assert result["slide_count"] == 5
    assert result["sections"][0] == {
        "section_index": 1,
        "name": "概要",
        "first_slide": 1,
        "slide_count": 2,
        "last_slide": 2,
        "section_id": "{SEC0001}",
    }
    assert result["sections"][1]["last_slide"] == 5


def test_sections_read_unsupported_without_section_properties() -> None:
    target = make_target([FakeSlide(1, "T1", "body")])
    with pytest.raises(BridgeError) as excinfo:
        powerpoint.sections(target)
    assert excinfo.value.code == ErrorCode.UNSUPPORTED_OPERATION
