from __future__ import annotations

from types import SimpleNamespace

import pytest

from office_bridge.errors import BridgeError, ErrorCode
from office_bridge.operations import word
from office_bridge.target import AppKind


class FakeSelection:
    def __init__(self, text: str = "Hello", start: int = 0, end: int = 5) -> None:
        self.Text = text
        self.Start = start
        self.End = end
        self.Type = 2
        self.typed = None
        self.collapsed = None

    def TypeText(self, text: str) -> None:
        self.typed = text

    def Collapse(self, direction: int) -> None:
        self.collapsed = direction


def make_target(selection: FakeSelection, document: str = "Report.docx") -> SimpleNamespace:
    return SimpleNamespace(
        app=AppKind.WORD,
        window=SimpleNamespace(Selection=selection),
        info={
            "application": "Word",
            "document": document,
            "full_name": rf"C:\docs\{document}",
        },
    )


def test_selection_reads_text_and_offsets() -> None:
    result = word.selection(make_target(FakeSelection()), max_text_chars=100)
    assert result["text"] == "Hello"
    assert result["characters"] == 5
    assert result["selection_type"] == "Normal"
    assert result["truncated"] is False


def test_selection_truncates_long_text() -> None:
    result = word.selection(
        make_target(FakeSelection("abcdefghij")), max_text_chars=4
    )
    assert result["text"] == "abcd"
    assert result["truncated"] is True


def test_insert_dry_run_does_not_change_selection() -> None:
    selection = FakeSelection()
    result = word.insert_text(
        make_target(selection),
        text="Replacement",
        replace=True,
        dry_run=True,
        max_text_chars=100,
        params={},
    )
    assert result["inserted"] is False
    assert selection.typed is None


def test_insert_requires_confirmation() -> None:
    selection = FakeSelection()
    with pytest.raises(BridgeError) as excinfo:
        word.insert_text(
            make_target(selection),
            text="Replacement",
            replace=True,
            dry_run=False,
            max_text_chars=100,
            params={},
        )
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED
    assert selection.typed is None


def test_insert_replaces_when_confirmed() -> None:
    selection = FakeSelection()
    result = word.insert_text(
        make_target(selection),
        text="Replacement",
        replace=True,
        dry_run=False,
        max_text_chars=100,
        params={"confirm": True},
    )
    assert result["inserted"] is True
    assert selection.typed == "Replacement"
    assert selection.collapsed is None


def test_insert_without_replace_collapses_to_start() -> None:
    selection = FakeSelection()
    word.insert_text(
        make_target(selection),
        text="Added",
        replace=False,
        dry_run=False,
        max_text_chars=100,
        params={"confirm": True},
    )
    assert selection.collapsed == 0
    assert selection.typed == "Added"


def test_insert_rejects_oversized_text() -> None:
    with pytest.raises(BridgeError) as excinfo:
        word.insert_text(
            make_target(FakeSelection()),
            text="abcdef",
            replace=True,
            dry_run=True,
            max_text_chars=3,
            params={},
        )
    assert excinfo.value.code == ErrorCode.TEXT_TOO_LARGE


# ---------------------------------------------------------------------------
# Phase 3B: paragraphs, find and replace
#
# The fake document stores paragraphs as plain strings. Reading a paragraph
# range's Text returns the text plus the trailing paragraph mark, assigning
# Text rewrites the paragraph (splitting on \r like Word does), and Find on a
# range performs literal replacements honouring MatchCase/MatchWholeWord and
# the wdReplaceOne/wdReplaceAll modes.
# ---------------------------------------------------------------------------


class FakeFind:
    def __init__(self, doc: "FakeWordDoc", indices: list[int]) -> None:
        self._doc = doc
        self._indices = indices
        self.Replacement = SimpleNamespace(ClearFormatting=lambda: None)
        self.executed: list[tuple] = []

    def ClearFormatting(self) -> None:
        pass

    def Execute(
        self,
        find_text,
        match_case,
        whole_words,
        match_wildcards,
        match_sounds_like,
        match_all_word_forms,
        forward,
        wrap,
        fmt,
        replace_with,
        replace_mode,
    ) -> bool:
        self.executed.append((find_text, replace_with, replace_mode))
        replaced = False
        for index in self._indices:
            text = self._doc.paragraphs[index]
            spans = word._match_spans(
                text, find_text, match_case=bool(match_case), whole_words=bool(whole_words)
            )
            if not spans:
                continue
            if replace_mode == word._WD_REPLACE_ONE:
                spans = spans[:1]
            for begin, finish in reversed(spans):
                text = text[:begin] + replace_with + text[finish:]
            self._doc.paragraphs[index] = text
            replaced = True
            if replace_mode == word._WD_REPLACE_ONE:
                break
        return replaced


class FakeParagraphRange:
    def __init__(self, doc: "FakeWordDoc", index: int) -> None:
        self._doc = doc
        self._index = index

    @property
    def Text(self) -> str:
        return self._doc.paragraphs[self._index] + "\r"

    @Text.setter
    def Text(self, value: str) -> None:
        text = value[:-1] if value.endswith("\r") else value
        self._doc.paragraphs[self._index : self._index + 1] = text.split("\r")

    @property
    def Find(self) -> FakeFind:
        return FakeFind(self._doc, [self._index])


class FakeContentRange:
    def __init__(self, doc: "FakeWordDoc") -> None:
        self._doc = doc

    @property
    def Find(self) -> FakeFind:
        return FakeFind(self._doc, list(range(len(self._doc.paragraphs))))


class FakeParagraph:
    def __init__(self, doc: "FakeWordDoc", index: int) -> None:
        self._doc = doc
        self._index = index

    @property
    def Range(self) -> FakeParagraphRange:
        return FakeParagraphRange(self._doc, self._index)

    @property
    def Style(self) -> str:
        return self._doc.styles.get(self._index, "Normal")


class FakeParagraphs:
    def __init__(self, doc: "FakeWordDoc") -> None:
        self._doc = doc

    @property
    def Count(self) -> int:
        return len(self._doc.paragraphs)

    def __call__(self, index: int) -> FakeParagraph:
        if index < 1 or index > self.Count:
            raise IndexError(f"paragraph {index} is out of range")
        return FakeParagraph(self._doc, index - 1)


class FakeWordDoc:
    def __init__(
        self,
        paragraphs: list[str],
        styles: dict[int, str] | None = None,
        track_revisions: bool = False,
    ) -> None:
        self.paragraphs = list(paragraphs)
        self.styles = dict(styles or {})
        self.TrackRevisions = track_revisions
        self.Paragraphs = FakeParagraphs(self)

    @property
    def Content(self) -> FakeContentRange:
        return FakeContentRange(self)


def make_doc_target(
    paragraphs: list[str],
    styles: dict[int, str] | None = None,
    track_revisions: bool = False,
    document: str = "Report.docx",
) -> SimpleNamespace:
    doc = FakeWordDoc(paragraphs, styles=styles, track_revisions=track_revisions)
    return SimpleNamespace(
        app=AppKind.WORD,
        window=SimpleNamespace(Selection=FakeSelection()),
        document=doc,
        info={
            "application": "Word",
            "document": document,
            "full_name": rf"C:\docs\{document}",
        },
    )


def test_paragraphs_reads_window_with_styles_and_headings() -> None:
    target = make_doc_target(
        ["Title", "Body one", "Body two", "Body three"],
        styles={0: "Title"},
    )
    result = word.paragraphs(target, start=2, count=2, max_text_chars=100)
    assert result["total_paragraphs"] == 4
    assert result["start"] == 2
    assert result["count"] == 2
    assert [item["text"] for item in result["paragraphs"]] == ["Body one", "Body two"]
    assert result["paragraphs"][0]["index"] == 2

    first = word.paragraphs(target, start=1, count=2, max_text_chars=100)["paragraphs"][0]
    assert first["style"] == "Title"
    assert first["is_heading"] is True
    assert first["characters"] == 5


def test_paragraphs_truncates_long_paragraphs() -> None:
    target = make_doc_target(["x" * 30])
    result = word.paragraphs(target, start=1, count=1, max_text_chars=10)
    assert result["paragraphs"][0]["text"] == "x" * 10
    assert result["paragraphs"][0]["truncated"] is True
    assert result["truncated"] is True


def test_paragraphs_rejects_bad_window() -> None:
    target = make_doc_target(["a", "b"])
    for bad in ((0, 10), ("x", 10), (1, 0), (1, 201), (1, -3)):
        with pytest.raises(BridgeError) as excinfo:
            word.paragraphs(target, start=bad[0], count=bad[1], max_text_chars=100)
        assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_find_reports_paragraph_numbers_and_context() -> None:
    target = make_doc_target(
        ["Alpha beta", "beta gamma beta", "nothing here", "BETA upper"]
    )
    result = word.find(
        target,
        find_text="beta",
        match_case=False,
        whole_words=False,
        max_matches=50,
        max_text_chars=200,
    )
    assert result["matches"] == 4
    assert result["more_matches"] is False
    assert [loc["paragraph"] for loc in result["locations"]] == [1, 2, 2, 4]
    assert "gamma" in result["locations"][1]["context"]


def test_find_respects_case_and_word_options_and_limit() -> None:
    target = make_doc_target(["Beta beta betabiosis", "beta"])
    cased = word.find(
        target,
        find_text="beta",
        match_case=True,
        whole_words=False,
        max_matches=50,
        max_text_chars=200,
    )
    assert cased["matches"] == 3  # "Beta" excluded, "betabiosis" included
    words = word.find(
        target,
        find_text="beta",
        match_case=False,
        whole_words=True,
        max_matches=50,
        max_text_chars=200,
    )
    assert words["matches"] == 3  # "betabiosis" excluded, "Beta" included
    limited = word.find(
        target,
        find_text="beta",
        match_case=False,
        whole_words=False,
        max_matches=2,
        max_text_chars=200,
    )
    assert limited["matches"] == 2
    assert limited["more_matches"] is True


def test_find_rejects_empty_needle() -> None:
    target = make_doc_target(["a"])
    for bad in ("", None, 3):
        with pytest.raises(BridgeError) as excinfo:
            word.find(
                target,
                find_text=bad,
                match_case=False,
                whole_words=False,
                max_matches=50,
                max_text_chars=200,
            )
        assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_replace_dry_run_reports_without_change() -> None:
    target = make_doc_target(["old value", "old again", "keep"])
    result = word.replace(
        target,
        find_text="old",
        replacement="new",
        match_case=False,
        whole_words=False,
        nth=None,
        dry_run=True,
        params={},
    )
    assert result["occurrences_before"] == 2
    assert result["first_paragraph"] == 1
    assert result["replaced"] is False
    assert target.document.paragraphs == ["old value", "old again", "keep"]


def test_replace_requires_confirmation() -> None:
    target = make_doc_target(["old"])
    with pytest.raises(BridgeError) as excinfo:
        word.replace(
            target,
            find_text="old",
            replacement="new",
            match_case=False,
            whole_words=False,
            nth=None,
            dry_run=False,
            params={},
        )
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED
    assert target.document.paragraphs == ["old"]


def test_replace_all_replaces_and_verifies() -> None:
    target = make_doc_target(["old value", "old again", "keep"])
    result = word.replace(
        target,
        find_text="old",
        replacement="new",
        match_case=False,
        whole_words=False,
        nth=None,
        dry_run=False,
        params={"confirm": True},
    )
    assert result["replaced"] is True
    assert result["verified"] is True
    assert result["occurrences_after"] == 0
    assert target.document.paragraphs == ["new value", "new again", "keep"]


def test_replace_nth_replaces_single_match() -> None:
    target = make_doc_target(["beta gamma beta", "beta end"])
    result = word.replace(
        target,
        find_text="beta",
        replacement="alpha",
        match_case=False,
        whole_words=False,
        nth=2,
        dry_run=False,
        params={"confirm": True},
    )
    assert result["nth"] == 2
    assert result["verified"] is True
    assert result["occurrences_after"] == 2
    # only the first match inside paragraph 1 changed
    assert target.document.paragraphs == ["alpha gamma beta", "beta end"]


def test_replace_nth_beyond_matches_rejected() -> None:
    target = make_doc_target(["only one"])
    with pytest.raises(BridgeError) as excinfo:
        word.replace(
            target,
            find_text="only",
            replacement="just",
            match_case=False,
            whole_words=False,
            nth=5,
            dry_run=True,
            params={},
        )
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST
    assert excinfo.value.details["occurrences"] == 1


def test_replace_notes_track_revisions() -> None:
    target = make_doc_target(["old"], track_revisions=True)
    result = word.replace(
        target,
        find_text="old",
        replacement="new",
        match_case=False,
        whole_words=False,
        nth=None,
        dry_run=True,
        params={},
    )
    assert result["track_revisions"] is True
    assert "tracked changes" in result["note"]


def test_replace_missing_needle_rejected() -> None:
    target = make_doc_target(["text"])
    with pytest.raises(BridgeError) as excinfo:
        word.replace(
            target,
            find_text="",
            replacement="new",
            match_case=False,
            whole_words=False,
            nth=None,
            dry_run=False,
            params={"confirm": True},
        )
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST
    with pytest.raises(BridgeError) as excinfo:
        word.replace(
            target,
            find_text="text",
            replacement=None,
            match_case=False,
            whole_words=False,
            nth=None,
            dry_run=False,
            params={"confirm": True},
        )
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST


def test_paragraph_replace_dry_run_previews() -> None:
    target = make_doc_target(["first", "second paragraph", "third"])
    result = word.replace_paragraph_text(
        target,
        paragraph=2,
        text="rewritten",
        dry_run=True,
        max_text_chars=100,
        params={},
    )
    assert result["text_preview"] == "second paragraph"
    assert result["paragraph_count_before"] == 3
    assert result["paragraph_count_after"] == 3
    assert result["replaced"] is False
    assert target.document.paragraphs[1] == "second paragraph"


def test_paragraph_replace_requires_confirmation() -> None:
    target = make_doc_target(["first"])
    with pytest.raises(BridgeError) as excinfo:
        word.replace_paragraph_text(
            target,
            paragraph=1,
            text="new",
            dry_run=False,
            max_text_chars=100,
            params={},
        )
    assert excinfo.value.code == ErrorCode.CONFIRMATION_REQUIRED
    assert target.document.paragraphs == ["first"]


def test_paragraph_replace_keeps_paragraph_break() -> None:
    target = make_doc_target(["first", "second", "third"])
    result = word.replace_paragraph_text(
        target,
        paragraph=2,
        text="rewritten",
        dry_run=False,
        max_text_chars=100,
        params={"confirm": True},
    )
    assert result["replaced"] is True
    assert result["verified"] is True
    assert target.document.paragraphs == ["first", "rewritten", "third"]
    assert result["paragraph_count_after"] == 3


def test_paragraph_replace_splits_on_newlines() -> None:
    target = make_doc_target(["first", "second", "third"])
    result = word.replace_paragraph_text(
        target,
        paragraph=2,
        text="line one\nline two",
        dry_run=False,
        max_text_chars=100,
        params={"confirm": True},
    )
    assert result["verified"] is True
    assert target.document.paragraphs == ["first", "line one", "line two", "third"]
    assert result["paragraph_count_after"] == 4


def test_paragraph_replace_rejects_bad_arguments() -> None:
    target = make_doc_target(["only"])
    for bad_paragraph in (0, "x", True):
        with pytest.raises(BridgeError) as excinfo:
            word.replace_paragraph_text(
                target,
                paragraph=bad_paragraph,
                text="new",
                dry_run=True,
                max_text_chars=100,
                params={},
            )
        assert excinfo.value.code == ErrorCode.INVALID_REQUEST
    with pytest.raises(BridgeError) as excinfo:
        word.replace_paragraph_text(
            target,
            paragraph=2,
            text="new",
            dry_run=True,
            max_text_chars=100,
            params={},
        )
    assert excinfo.value.code == ErrorCode.INVALID_REQUEST
    with pytest.raises(BridgeError) as excinfo:
        word.replace_paragraph_text(
            target,
            paragraph=1,
            text="way too long",
            dry_run=True,
            max_text_chars=5,
            params={},
        )
    assert excinfo.value.code == ErrorCode.TEXT_TOO_LARGE
