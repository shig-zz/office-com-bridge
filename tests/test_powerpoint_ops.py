from __future__ import annotations

from types import SimpleNamespace

import pytest

from office_bridge.errors import BridgeError, ErrorCode
from office_bridge.operations import powerpoint
from office_bridge.target import AppKind


class FakeTextRange:
    def __init__(self, text: str) -> None:
        self.Text = text


class FakeTextFrame:
    def __init__(self, text: str) -> None:
        self.HasText = bool(text)
        self.TextRange = FakeTextRange(text)


class FakeShape:
    def __init__(self, shape_id: int, name: str, text: str) -> None:
        self.Id = shape_id
        self.Name = name
        self.Type = 14
        self.HasTextFrame = True
        self.TextFrame = FakeTextFrame(text)
        self.HasTable = False


class FakeShapes:
    def __init__(self, shapes: list[FakeShape], title: FakeShape | None) -> None:
        self._shapes = shapes
        self.Count = len(shapes)
        self.Title = title

    def __call__(self, index: int) -> FakeShape:
        return self._shapes[index - 1]


class FakeSlide:
    def __init__(self, index: int, title: str, body: str) -> None:
        title_shape = FakeShape(1, "Title", title)
        body_shape = FakeShape(2, "Body", body)
        self.SlideIndex = index
        self.SlideID = 200 + index
        self.CustomLayout = SimpleNamespace(Name="Title and Content")
        self.Shapes = FakeShapes([title_shape, body_shape], title_shape)


class FakeSlides:
    def __init__(self, slides: list[FakeSlide]) -> None:
        self._slides = slides
        self.Count = len(slides)

    def __call__(self, index: int) -> FakeSlide:
        return self._slides[index - 1]


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
