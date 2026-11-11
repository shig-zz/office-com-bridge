from __future__ import annotations

import re
from typing import Any, Callable

import pywintypes

from ..errors import (
    BridgeError,
    ErrorCode,
    invalid_request,
    range_too_large,
)
from ..target import ResolvedTarget
from .common import (
    check_cell_limit,
    check_expected_document,
    jsonable,
    optional_bool,
    require_confirm,
    require_str,
    validate_matrix,
)

# Enumerated values verified against the installed Excel type library
# (Microsoft.Office.Interop.Excel).
_XL_PASTE_FORMATS = -4122
_XL_HALIGN = {"left": -4131, "center": -4108, "right": -4152}
_XL_VALIGN = {"top": -4160, "middle": -4108, "bottom": -4107}
_XL_LINE_CONTINUOUS = 1
_XL_PATTERN_SOLID = 1
_XL_PATTERN_NONE = -4142
_XL_INSERT_SHIFT = {"down": -4107, "right": -4161}
_XL_DELETE_SHIFT = {"up": -4122, "left": -4121}
_XL_BORDER_SIDES = {"left": 7, "top": 8, "bottom": 9, "right": 10}
_XL_LINE_STYLES = {
    -4142: "none",
    1: "continuous",
    3: "dotted",
    4: "dashed",
    5: "dash dot",
    6: "dash dot dot",
    7: "double",
    13: "slant dash dot",
}


def _dims(range_object: Any) -> tuple[int, int]:
    try:
        return int(range_object.Rows.Count), int(range_object.Columns.Count)
    except pywintypes.com_error as exc:
        raise _range_error(exc) from exc


def _parse_color(value: Any, key: str) -> int:
    if not isinstance(value, str):
        raise invalid_request(
            f"Parameter {key!r} must be a 'RRGGBB' or '#RRGGBB' hex color string.",
            {"parameter": key},
        )
    text = value.strip().lstrip("#")
    if len(text) != 6:
        raise invalid_request(
            f"Parameter {key!r} must be a 'RRGGBB' or '#RRGGBB' hex color string.",
            {"parameter": key},
        )
    try:
        rgb = int(text, 16)
    except ValueError:
        raise invalid_request(
            f"Parameter {key!r} is not a valid hex color: {value!r}.",
            {"parameter": key},
        ) from None
    red = (rgb >> 16) & 0xFF
    green = (rgb >> 8) & 0xFF
    blue = rgb & 0xFF
    # Excel stores colors as BGR.
    return blue * 0x10000 + green * 0x100 + red


def _set_color_property(range_object: Any, target: str, color: int) -> None:
    if target == "Font.Color":
        range_object.Font.Color = color
    elif target == "Interior.Color":
        range_object.Interior.Pattern = _XL_PATTERN_SOLID
        range_object.Interior.Color = color
    else:
        range_object.Borders.Color = color


def _address(range_object: Any) -> str:
    try:
        value = range_object.Address(False, False)
    except Exception:
        value = range_object.Address
    return str(value).replace("$", "")


def _matrix(value: Any, rows: int, columns: int) -> list[list[Any]]:
    if rows == 1 and columns == 1:
        return [[jsonable(value)]]
    if value is None:
        return [[None for _ in range(columns)] for _ in range(rows)]

    if isinstance(value, (list, tuple)):
        if rows == 1:
            if len(value) == 1 and isinstance(value[0], (list, tuple)):
                row = list(value[0])
            else:
                row = list(value)
            row = [jsonable(cell) for cell in row]
            return [row]
        result: list[list[Any]] = []
        for item in value:
            if isinstance(item, (list, tuple)):
                result.append([jsonable(cell) for cell in item])
            else:
                result.append([jsonable(item)])
        return result
    return [[jsonable(value)]]


def _rectangular(matrix: list[list[Any]], rows: int, columns: int) -> list[list[Any]]:
    if len(matrix) == rows and all(len(row) == columns for row in matrix):
        return matrix
    result = [[None for _ in range(columns)] for _ in range(rows)]
    for row_index in range(min(rows, len(matrix))):
        for column_index in range(min(columns, len(matrix[row_index]))):
            result[row_index][column_index] = matrix[row_index][column_index]
    return result


def _read_area(range_object: Any, *, include_formulas: bool, max_cells: int) -> dict[str, Any]:
    try:
        rows = int(range_object.Rows.Count)
        columns = int(range_object.Columns.Count)
    except pywintypes.com_error as exc:
        raise _range_error(exc) from exc
    cells = rows * columns
    check_cell_limit(cells, max_cells)

    try:
        values = _rectangular(_matrix(range_object.Value2, rows, columns), rows, columns)
    except pywintypes.com_error as exc:
        raise _range_error(exc) from exc

    result: dict[str, Any] = {
        "address": _address(range_object),
        "rows": rows,
        "columns": columns,
        "cell_count": cells,
        "values": values,
    }
    if include_formulas:
        try:
            result["formulas"] = _rectangular(
                _matrix(range_object.Formula, rows, columns), rows, columns
            )
        except pywintypes.com_error as exc:
            raise _range_error(exc) from exc
    return result


def _range_error(exc: pywintypes.com_error) -> BridgeError:
    hresult = getattr(exc, "hresult", None)
    return BridgeError(
        ErrorCode.RANGE_INVALID,
        "The Excel range could not be read or written. Check the address and the active sheet.",
        400,
        {"error": str(exc), "hresult": f"0x{int(hresult) & 0xFFFFFFFF:08X}"} if hresult is not None else {"error": str(exc)},
    )


def _range(target: ResolvedTarget, address: str) -> Any:
    try:
        sheet = target.window.ActiveSheet
        return sheet.Range(address)
    except pywintypes.com_error as exc:
        raise _range_error(exc) from exc
    except Exception as exc:
        raise invalid_request(
            f"Could not resolve the Excel range {address!r}.",
            {"address": address},
        ) from exc


def read_range(
    target: ResolvedTarget,
    *,
    address: str,
    include_formulas: bool,
    max_cells: int,
) -> dict[str, Any]:
    range_object = _range(target, address)
    try:
        areas_count = int(range_object.Areas.Count)
    except Exception:
        areas_count = 1

    if areas_count > 1:
        areas: list[dict[str, Any]] = []
        total_cells = 0
        for index in range(1, areas_count + 1):
            area = range_object.Areas(index)
            try:
                total_cells += int(area.Count)
            except pywintypes.com_error as exc:
                raise _range_error(exc) from exc
            if total_cells > max_cells:
                raise range_too_large(total_cells, max_cells)
            areas.append(
                _read_area(area, include_formulas=include_formulas, max_cells=max_cells)
            )
        return {
            "address": _address(range_object),
            "multi_area": True,
            "cell_count": total_cells,
            "areas": areas,
        }

    result = _read_area(range_object, include_formulas=include_formulas, max_cells=max_cells)
    result["multi_area"] = False
    return result


def selection(target: ResolvedTarget, *, limit: int) -> dict[str, Any]:
    try:
        selected = target.window.RangeSelection
    except pywintypes.com_error as exc:
        raise _range_error(exc) from exc
    try:
        areas_count = int(selected.Areas.Count)
    except Exception:
        areas_count = 1
    if areas_count > 1:
        result = read_range(
            target,
            address=_address(selected),
            include_formulas=False,
            max_cells=limit,
        )
        result["selection_type"] = "Range"
        return result
    result = _read_area(selected, include_formulas=False, max_cells=limit)
    result["multi_area"] = False
    result["selection_type"] = "Range"
    return result


def document(target: ResolvedTarget) -> dict[str, Any]:
    workbook = target.document
    sheets: list[str] = []
    try:
        for index in range(1, int(workbook.Worksheets.Count) + 1):
            sheets.append(str(workbook.Worksheets(index).Name))
    except pywintypes.com_error:
        sheets = []
    info = {
        "application": "Excel",
        "name": target.info.get("document", ""),
        "full_name": target.info.get("full_name", ""),
        "path": target.info.get("path", ""),
        "active_sheet": target.info.get("sheet", ""),
        "sheets": sheets,
        "read_only": False,
    }
    try:
        info["read_only"] = bool(workbook.ReadOnly)
    except Exception:
        pass
    return info


def write_range(
    target: ResolvedTarget,
    *,
    address: str,
    values: Any,
    dry_run: bool,
    max_cells: int,
    params: dict[str, Any],
) -> dict[str, Any]:
    check_expected_document(target, params)
    matrix, rows, columns = validate_matrix(values)
    cells = rows * columns
    check_cell_limit(cells, max_cells)

    for row in matrix:
        for cell in row:
            if isinstance(cell, str) and cell.startswith("="):
                raise invalid_request(
                    "Formula strings beginning with '=' are not supported by excel.range.write. "
                    "Write ordinary values only.",
                    {"address": address},
                )

    range_object = _range(target, address)
    try:
        actual_rows = int(range_object.Rows.Count)
        actual_columns = int(range_object.Columns.Count)
    except pywintypes.com_error as exc:
        raise _range_error(exc) from exc

    if (actual_rows, actual_columns) != (rows, columns):
        raise BridgeError(
            ErrorCode.SIZE_MISMATCH,
            f"Range {address!r} is {actual_rows}x{actual_columns} but the supplied array is {rows}x{columns}.",
            400,
            {
                "address": address,
                "range_rows": actual_rows,
                "range_columns": actual_columns,
                "array_rows": rows,
                "array_columns": columns,
            },
        )

    if not dry_run:
        require_confirm(params, "excel.range.write")

    result = {
        "address": _address(range_object),
        "rows": rows,
        "columns": columns,
        "cell_count": cells,
        "dry_run": dry_run,
        "written": False,
    }
    if dry_run:
        return result

    try:
        if rows == 1 and columns == 1:
            range_object.Value2 = matrix[0][0]
        else:
            range_object.Value2 = tuple(tuple(row) for row in matrix)
    except pywintypes.com_error as exc:
        raise _range_error(exc) from exc

    result["written"] = True
    return result


def copy_format(
    target: ResolvedTarget,
    *,
    source_address: str,
    address: str,
    dry_run: bool,
    max_cells: int,
    params: dict[str, Any],
    include_borders: bool = False,
) -> dict[str, Any]:
    check_expected_document(target, params)
    source = _range(target, source_address)
    dest = _range(target, address)
    source_dims = _dims(source)
    dest_dims = _dims(dest)
    if source_dims != dest_dims:
        raise BridgeError(
            ErrorCode.SIZE_MISMATCH,
            f"Source range {source_address!r} is {source_dims[0]}x{source_dims[1]} but target "
            f"range {address!r} is {dest_dims[0]}x{dest_dims[1]}.",
            400,
            {
                "source_address": source_address,
                "address": address,
                "source": list(source_dims),
                "target": list(dest_dims),
            },
        )
    rows, columns = dest_dims
    check_cell_limit(rows * columns, max_cells)

    if not dry_run:
        require_confirm(params, "excel.range.copy_format")

    result: dict[str, Any] = {
        "source_address": _address(source),
        "address": _address(dest),
        "rows": rows,
        "columns": columns,
        "cell_count": rows * columns,
        "dry_run": dry_run,
        "copied": False,
    }
    if dry_run:
        return result

    try:
        source.Copy()
        dest.PasteSpecial(_XL_PASTE_FORMATS)
    except pywintypes.com_error as exc:
        try:
            target.application.CutCopyMode = False
        except Exception:
            pass
        raise _range_error(exc) from exc
    try:
        target.application.CutCopyMode = False
    except Exception:
        pass
    result["copied"] = True
    result["borders_applied"] = False
    if include_borders:
        # This build does not propagate cell borders through a format paste,
        # so re-apply them cell by cell from the source.
        try:
            for row in range(1, rows + 1):
                for column in range(1, columns + 1):
                    source_border_owner = source.Cells(row, column)
                    dest_border_owner = dest.Cells(row, column)
                    for position in _XL_BORDER_SIDES.values():
                        source_border = source_border_owner.Borders(position)
                        style = int(source_border.LineStyle)
                        dest_border = dest_border_owner.Borders(position)
                        dest_border.LineStyle = style
                        if style != _XL_PATTERN_NONE:
                            dest_border.Weight = int(source_border.Weight)
                            dest_border.Color = int(source_border.Color)
            result["borders_applied"] = True
        except pywintypes.com_error as exc:
            raise _range_error(exc) from exc
    else:
        result["borders_note"] = (
            "this build does not propagate cell borders through a format paste; "
            "run again with borders: true to copy them per cell"
        )
    return result


def set_format(
    target: ResolvedTarget,
    *,
    address: str,
    dry_run: bool,
    max_cells: int,
    params: dict[str, Any],
) -> dict[str, Any]:
    check_expected_document(target, params)
    range_object = _range(target, address)
    rows, columns = _dims(range_object)
    check_cell_limit(rows * columns, max_cells)

    planned: list[str] = []
    actions: list[Callable[[], None]] = []

    def add_font(attr: str, value: Any, label: str) -> None:
        actions.append(lambda attr=attr, value=value: setattr(range_object.Font, attr, value))
        planned.append(f"{label}={value!r}")

    if params.get("font_name_far_east") is not None:
        raise invalid_request(
            "Excel's Font object has no settable NameFarEast property. "
            "Use 'font_name' with a pan-CJK font (for example 'MS Gothic', 'Yu Gothic', "
            "or 'Meiryo'); Excel applies it to both Japanese and Western text.",
            {"parameter": "font_name_far_east"},
        )

    font_name = params.get("font_name")
    if font_name is not None:
        add_font("Name", require_str(params, "font_name"), "Font.Name")

    font_size = params.get("font_size")
    if font_size is not None:
        if isinstance(font_size, bool) or not isinstance(font_size, (int, float)) or not 0 < float(font_size) <= 1000:
            raise invalid_request("Parameter 'font_size' must be a positive number.", {"parameter": "font_size"})
        add_font("Size", float(font_size), "Font.Size")

    for key, attr in (("bold", "Bold"), ("italic", "Italic"), ("underline", "Underline")):
        value = params.get(key)
        if value is not None:
            if not isinstance(value, bool):
                raise invalid_request(f"Parameter {key!r} must be a boolean.", {"parameter": key})
            add_font(attr, value, f"Font.{attr}")

    number_format = params.get("number_format")
    if number_format is not None:
        number_format = require_str(params, "number_format")
        actions.append(lambda: setattr(range_object, "NumberFormat", number_format))
        planned.append(f"NumberFormat={number_format!r}")

    halign = params.get("halign")
    if halign is not None:
        if not isinstance(halign, str) or halign.lower() not in _XL_HALIGN:
            raise invalid_request("Parameter 'halign' must be one of: left, center, right.", {"parameter": "halign"})
        mapped = _XL_HALIGN[halign.lower()]
        actions.append(lambda: setattr(range_object, "HorizontalAlignment", mapped))
        planned.append(f"HorizontalAlignment={halign!r}")

    valign = params.get("valign")
    if valign is not None:
        if not isinstance(valign, str) or valign.lower() not in _XL_VALIGN:
            raise invalid_request("Parameter 'valign' must be one of: top, middle, bottom.", {"parameter": "valign"})
        mapped = _XL_VALIGN[valign.lower()]
        actions.append(lambda: setattr(range_object, "VerticalAlignment", mapped))
        planned.append(f"VerticalAlignment={valign!r}")

    wrap = params.get("wrap")
    if wrap is not None:
        if not isinstance(wrap, bool):
            raise invalid_request("Parameter 'wrap' must be a boolean.", {"parameter": "wrap"})
        actions.append(lambda: setattr(range_object, "WrapText", wrap))
        planned.append(f"WrapText={wrap!r}")

    for key, color_target in (("color", "Font.Color"), ("fill", "Interior.Color"), ("border_color", "Borders.Color")):
        value = params.get(key)
        if value is not None:
            parsed = _parse_color(value, key)
            actions.append(lambda color_target=color_target, parsed=parsed: _set_color_property(range_object, color_target, parsed))
            planned.append(f"{color_target}={value.strip().lstrip('#').upper()}")

    border = params.get("border")
    if border is not None:
        if not isinstance(border, bool) and not (isinstance(border, str) and border.strip().lower() == "continuous"):
            raise invalid_request("Parameter 'border' must be true or 'continuous'.", {"parameter": "border"})
        actions.append(lambda: setattr(range_object.Borders, "LineStyle", _XL_LINE_CONTINUOUS))
        planned.append("Borders.LineStyle=continuous")

    if not actions:
        raise invalid_request(
            "Provide at least one formatting option: font_name, font_size, bold, "
            "italic, underline, color, number_format, halign, valign, wrap, fill, border, border_color.",
            {"parameter": "address"},
        )

    if not dry_run:
        require_confirm(params, "excel.range.set_format")
        try:
            for action in actions:
                action()
        except pywintypes.com_error as exc:
            raise _range_error(exc) from exc

    return {
        "address": _address(range_object),
        "rows": rows,
        "columns": columns,
        "cell_count": rows * columns,
        "dry_run": dry_run,
        "applied": not dry_run,
        "changes": planned,
    }


def add_hyperlink(
    target: ResolvedTarget,
    *,
    address: str,
    dry_run: bool,
    max_cells: int,
    params: dict[str, Any],
) -> dict[str, Any]:
    check_expected_document(target, params)
    url = require_str(params, "url")
    if url.startswith("="):
        raise invalid_request(
            "Hyperlink addresses beginning with '=' are not supported by "
            "excel.range.add_hyperlink. Provide a URL, or use 'sub_address' for a "
            "sheet reference such as 'Sheet2!A1'.",
            {"address": address},
        )
    screen_tip = params.get("screen_tip")
    if screen_tip is not None and (
        not isinstance(screen_tip, str) or not screen_tip.strip()
    ):
        raise invalid_request(
            "Parameter 'screen_tip' must be a non-empty string.",
            {"parameter": "screen_tip"},
        )
    sub_address = params.get("sub_address")
    if sub_address is not None and (not isinstance(sub_address, str) or not sub_address.strip()):
        raise invalid_request(
            "Parameter 'sub_address' must be a non-empty string.",
            {"parameter": "sub_address"},
        )
    if sub_address is not None and sub_address.startswith("="):
        raise invalid_request(
            "Sub addresses beginning with '=' are not supported by "
            "excel.range.add_hyperlink. Use a sheet reference such as 'Sheet2!A1' "
            "without a leading '='.",
            {"parameter": "sub_address"},
        )

    range_object = _range(target, address)
    rows, columns = _dims(range_object)
    check_cell_limit(rows * columns, max_cells)

    try:
        link_count = int(range_object.Hyperlinks.Count)
    except pywintypes.com_error as exc:
        raise _range_error(exc) from exc

    if not dry_run:
        require_confirm(params, "excel.range.add_hyperlink")

    result: dict[str, Any] = {
        "address": _address(range_object),
        "rows": rows,
        "columns": columns,
        "cell_count": rows * columns,
        "url": url,
        "action": "update" if link_count > 0 else "add",
        "existing_links": link_count,
        "dry_run": dry_run,
        "applied": False,
    }
    if dry_run:
        return result

    try:
        if link_count > 0:
            link = range_object.Hyperlinks(1)
            link.Address = url
            if sub_address is not None:
                link.SubAddress = sub_address
            if screen_tip is not None:
                link.ScreenTip = screen_tip
        elif sub_address is not None and screen_tip is not None:
            range_object.Hyperlinks.Add(range_object, url, sub_address, screen_tip)
        elif sub_address is not None:
            range_object.Hyperlinks.Add(range_object, url, sub_address)
        elif screen_tip is not None:
            range_object.Hyperlinks.Add(range_object, url, "", screen_tip)
        else:
            range_object.Hyperlinks.Add(range_object, url)
    except pywintypes.com_error as exc:
        raise _range_error(exc) from exc

    result["applied"] = True
    return result


_XL_SHIFT_DOWN = -4121


def _column_letter(index: int) -> str:
    letter = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        letter = chr(65 + rem) + letter
    return letter


def _letter_column(letter: str) -> int:
    index = 0
    for char in letter:
        index = index * 26 + (ord(char) - 64)
    return index


def _parse_sort_key(raw: Any, column_count: int, start_column: int) -> tuple[int, bool]:
    """Parse one sort key entry.

    Returns (column_index, ascending) where column_index is 1-based from the left edge
    of the sort range. The key column may be given as that 1-based offset or as a sheet
    column letter (for example "C") that must lie inside the sort range.
    """
    if not isinstance(raw, dict):
        raise invalid_request(
            "Each sort key must be an object like {\"column\": 1, \"order\": \"ascending\"}.",
            {"parameter": "keys"},
        )
    col = raw.get("column")
    if isinstance(col, bool):
        raise invalid_request(
            "Sort key 'column' must be a 1-based column offset number or a sheet column "
            "letter such as 'C'.",
            {"parameter": "keys"},
        )
    if isinstance(col, int):
        column_index = col
    elif isinstance(col, str):
        letter = col.strip().upper().lstrip("$")
        if not letter or not letter.isalpha() or len(letter) > 3:
            raise invalid_request(
                f"Sort key column {col!r} is not a column letter or number.",
                {"parameter": "keys"},
            )
        column_index = _letter_column(letter) - start_column + 1
    else:
        raise invalid_request(
            "Sort key 'column' must be a 1-based column offset number or a sheet column "
            "letter such as 'C'.",
            {"parameter": "keys"},
        )
    if not 1 <= column_index <= column_count:
        raise invalid_request(
            f"Sort key column is outside the sort range (columns 1-{column_count} from its left edge).",
            {"parameter": "keys"},
        )
    order = raw.get("order", "ascending")
    if order == "ascending":
        return column_index, True
    if order == "descending":
        return column_index, False
    raise invalid_request(
        "Sort key 'order' must be 'ascending' or 'descending'.",
        {"parameter": "keys"},
    )


def _sort_key_value(value: Any) -> tuple[int, float, str]:
    """Total ordering for mixed cell values: blanks, then numbers, then text, then others."""
    if value is None:
        return (0, 0.0, "")
    if isinstance(value, bool):
        return (1, 1.0 if value else 0.0, "")
    if isinstance(value, (int, float)):
        return (1, float(value), "")
    if isinstance(value, str):
        return (2, 0.0, value)
    return (3, 0.0, str(value))


def _compute_order(values: list[list[Any]], keys: list[tuple[int, bool]]) -> list[int]:
    """Stable multi-key sort. Returns the original 0-based row indices in target order.

    keys is applied in order (first key is primary). Later keys only break ties of the
    earlier ones, matching Excel's multi-level sort.
    """
    order = list(range(len(values)))
    for column_index, ascending in reversed(keys):
        order = sorted(
            order,
            key=lambda row_index: _sort_key_value(values[row_index][column_index - 1]),
            reverse=not ascending,
        )
    return order


def _moves_for(order: list[int]) -> list[tuple[int, int]]:
    """Row moves that realize the target order.

    Each move is (destination, source) with 1-based positions, and source is always below
    destination: the row at source is lifted into destination, shifting the rows between.
    """
    current = list(range(len(order)))
    moves: list[tuple[int, int]] = []
    for position, wanted in enumerate(order):
        if current[position] == wanted:
            continue
        source = current.index(wanted)
        moves.append((position + 1, source + 1))
        current = current[:position] + [wanted] + current[position:source] + current[source + 1 :]
    return moves


def sort_range(
    target: ResolvedTarget,
    *,
    address: str,
    keys: Any,
    dry_run: bool,
    max_cells: int,
    params: dict[str, Any],
) -> dict[str, Any]:
    """Sort the rows of a range in place.

    This Excel build silently ignores Range.Sort calls whose key is a Range object passed
    through automation (the call returns success but nothing is reordered), and the
    ListObject sort path is restricted (only the column-orientation value is accepted).
    Rows are therefore moved with the primitive row operations that do work: the whole
    source row is copied, an insert at the destination pastes the copy (the inserted row
    receives the clipboard contents in this build), and the now-duplicated source row is
    deleted. The copy must span the whole row: inserting a whole row with a narrower
    clipboard range is rejected by this Excel build as an overlapping copy/paste area.
    Values, cell formatting, hyperlinks and row structure travel with the row.
    """
    check_expected_document(target, params)
    if not isinstance(keys, list) or not 1 <= len(keys) <= 3:
        raise invalid_request(
            "Parameter 'keys' must be an array of 1 to 3 sort key objects.",
            {"parameter": "keys"},
        )

    range_object = _range(target, address)
    try:
        rows, columns = _dims(range_object)
        top_row = int(range_object.Row)
        start_column = int(range_object.Columns(1).Column)
    except pywintypes.com_error as exc:
        raise _range_error(exc) from exc
    check_cell_limit(rows * columns, max_cells)
    if rows < 2:
        raise invalid_request("The sort range needs at least two rows.", {"address": address})

    try:
        values = _rectangular(_matrix(range_object.Value2, rows, columns), rows, columns)
    except pywintypes.com_error as exc:
        raise _range_error(exc) from exc

    parsed_keys = [_parse_sort_key(key, columns, start_column) for key in keys]
    order = _compute_order(values, parsed_keys)
    moves = _moves_for(order)

    result: dict[str, Any] = {
        "address": _address(range_object),
        "rows": rows,
        "columns": columns,
        "keys": [
            {"column": index, "order": "ascending" if ascending else "descending"}
            for index, ascending in parsed_keys
        ],
        "moves": len(moves),
        "row_order": [row_index + 1 for row_index in order],
        "dry_run": dry_run,
        "applied": False,
    }
    if dry_run:
        return result

    require_confirm(params, "excel.range.sort")

    sheet = range_object.Parent
    try:
        # Read heights per row: Range.RowHeight is unreliable in this build (it can
        # return None or a scalar for multi-row ranges), while Rows(n).RowHeight works.
        heights: list[float | None] = []
        for position in range(rows):
            height_value = sheet.Rows(top_row + position).RowHeight
            heights.append(float(height_value) if height_value is not None else None)
        known_heights = [height for height in heights if height is not None]
        uniform_height = len(known_heights) > 0 and len(set(known_heights)) <= 1
        all_heights_known = len(known_heights) == rows

        for destination, source in moves:
            source_row = top_row + source - 1
            destination_row = top_row + destination - 1
            sheet.Rows(source_row).Copy()
            sheet.Rows(destination_row).Insert(_XL_SHIFT_DOWN)
            target.application.CutCopyMode = False
            sheet.Rows(source_row + 1).Delete()

        if not uniform_height and all_heights_known:
            # The inserted rows inherit a neighbour's height, so re-apply the heights to
            # follow the new row order. Only done when every height is known: assigning
            # None to RowHeight raises a COM error, and skipping would lose the heights
            # of the known rows anyway.
            for position, row_index in enumerate(order):
                height = heights[row_index]
                if height is not None:
                    sheet.Rows(top_row + position).RowHeight = height

        # Re-resolve the range for the verification read: in this build a Range object
        # obtained before the row moves shrinks as rows are deleted, so reading through
        # it would return fewer rows than the range actually holds.
        final_range = _range(target, address)
        final_values = _rectangular(_matrix(final_range.Value2, rows, columns), rows, columns)
    except pywintypes.com_error as exc:
        raise _range_error(exc) from exc

    expected = [values[row_index] for row_index in order]
    result["applied"] = True
    result["verified"] = final_values == expected
    return result


def set_active_sheet(target: ResolvedTarget, value: Any, params: dict[str, Any]) -> dict[str, Any]:
    if not isinstance(value, str) or not value.strip():
        raise invalid_request("The value for excel.active_sheet must be a sheet name.")
    require_confirm(params, "excel.active_sheet")
    check_expected_document(target, params)
    workbook = target.document
    sheet = None
    try:
        for index in range(1, int(workbook.Worksheets.Count) + 1):
            candidate = workbook.Worksheets(index)
            if str(candidate.Name).lower() == value.lower():
                sheet = candidate
                break
    except pywintypes.com_error as exc:
        raise _range_error(exc) from exc
    if sheet is None:
        raise invalid_request(f"Worksheet {value!r} was not found.", {"sheet": value})
    try:
        sheet.Activate()
    except pywintypes.com_error as exc:
        raise _range_error(exc) from exc
    target.refresh_info()
    return {"active_sheet": value}


# ---------------------------------------------------------------------------
# Structural writes: insert/delete whole rows or columns, move a row
#
# Whole-row and whole-column insert/delete work through COM on this build
# (unlike Range.Sort), but the shift enums must be exact and every range
# reference must be re-resolved afterwards. These operations wrap both and
# verify through the sheet's used range.
# ---------------------------------------------------------------------------

_ROW_SPEC = re.compile(r"^\d+(?::\d+)?$")
_COLUMN_SPEC = re.compile(r"^\d+(?::\d+)?$|^[A-Za-z]{1,3}(?::[A-Za-z]{1,3})?$")
_DIGIT_RANGE = re.compile(r"^\d+:\d+$")
_UNDO_HINT = (
    "roll this back with the 'undo' operation only after the user approves; "
    "it reverts the most recent Office action"
)


def _sheet_name(sheet: Any) -> str:
    try:
        return str(sheet.Name)
    except Exception:
        return ""


def _last_used_index(sheet: Any, dimension: str) -> int | None:
    """Last row (or column) inside the sheet's used range, None if unreadable."""
    try:
        used = sheet.UsedRange
        if dimension == "rows":
            first = int(used.Row)
            count = int(used.Rows.Count)
        else:
            first = int(used.Column)
            count = int(used.Columns.Count)
    except Exception:
        return None
    if count <= 0:
        return None
    return first + count - 1


def _validate_block_spec(spec: Any, parameter: str, dimension: str) -> Any:
    if isinstance(spec, bool) or not isinstance(spec, (int, str)):
        raise invalid_request(
            f"Parameter {parameter!r} must be a number or a range string.",
            {"parameter": parameter},
        )
    text = str(spec).strip()
    if not (_ROW_SPEC if dimension == "rows" else _COLUMN_SPEC).match(text):
        example = '"61" or "2:4"' if dimension == "rows" else '"C", "3", or "A:C"'
        raise invalid_request(
            f"Parameter {parameter!r} must be a {dimension[:-1]} number or range "
            f"such as {example}.",
            {"parameter": parameter, parameter: spec},
        )
    if text.isdigit() or _DIGIT_RANGE.match(text):
        parts = [int(part) for part in text.split(":")]
        if any(part < 1 for part in parts) or (len(parts) == 2 and parts[1] < parts[0]):
            raise invalid_request(
                f"Parameter {parameter!r} must name at least one {dimension[:-1]} "
                "with a positive number.",
                {"parameter": parameter, parameter: spec},
            )
        return int(text) if ":" not in text else text
    return text.upper()


def insert_or_delete(
    target: ResolvedTarget,
    *,
    dimension: str,
    spec: Any,
    shift: Any,
    mode: str,
    dry_run: bool,
    params: dict[str, Any],
) -> dict[str, Any]:
    """Insert or delete whole rows/columns on the active sheet."""
    check_expected_document(target, params)
    if dimension not in ("rows", "columns") or mode not in ("insert", "delete"):
        raise invalid_request(
            "The dimension must be 'rows' or 'columns' and the mode 'insert' or 'delete'.",
            {"dimension": dimension, "mode": mode},
        )
    parameter = "row" if dimension == "rows" else "column"
    resolved = _validate_block_spec(spec, parameter, dimension)
    shift_table = _XL_INSERT_SHIFT if mode == "insert" else _XL_DELETE_SHIFT
    if shift is None:
        shift_name = None
    else:
        if not isinstance(shift, str) or shift.lower() not in shift_table:
            raise invalid_request(
                f"Parameter 'shift' must be one of: {', '.join(shift_table)}.",
                {"parameter": "shift"},
            )
        shift_name = shift.lower()
    default_shift = (
        "down"
        if mode == "insert" and dimension == "rows"
        else "right"
        if mode == "insert"
        else "up"
        if dimension == "rows"
        else "left"
    )
    try:
        sheet = target.window.ActiveSheet
        picker = sheet.Rows if dimension == "rows" else sheet.Columns
        block = picker(resolved)
        count = int(block.Rows.Count) if dimension == "rows" else int(block.Columns.Count)
        first = int(block.Row) if dimension == "rows" else int(block.Column)
    except BridgeError:
        raise
    except pywintypes.com_error as exc:
        raise _range_error(exc) from exc
    except Exception as exc:
        raise invalid_request(
            f"Could not resolve {dimension} {spec!r} on the active sheet.",
            {"parameter": parameter},
        ) from exc
    last_before = _last_used_index(sheet, dimension)
    result: dict[str, Any] = {
        "sheet": _sheet_name(sheet),
        parameter: resolved,
        "first": first,
        "count": count,
        "shift": shift_name or default_shift,
        "last_used_before": last_before,
        "dry_run": dry_run,
        "applied": False,
    }
    if dry_run:
        return result
    require_confirm(params, f"excel.{dimension}.{mode}")
    try:
        if mode == "insert":
            block.Insert(shift_table[shift_name or default_shift])
        elif shift_name is None:
            # Rows(n).Delete() with no shift argument is the reliable form on
            # this build; a shift is only passed when the caller chose one.
            block.Delete()
        else:
            block.Delete(shift_table[shift_name])
    except pywintypes.com_error as exc:
        raise _range_error(exc) from exc
    last_after = _last_used_index(sheet, dimension)
    if last_before is None or last_after is None:
        verified: bool | None = None
    elif mode == "insert":
        if first <= last_before:
            # Inserting inside the used area pushes the end down by count.
            expected = last_before + count
        else:
            # Inserting below the used area extends the used range to the
            # inserted rows themselves (blank inserted rows still count as
            # used on this build).
            expected = max(last_before, first + count - 1)
        verified = last_after == expected
    else:
        removed = 0 if first > last_before else min(count, last_before - first + 1)
        verified = last_after == last_before - removed
    result.update(
        {
            "applied": True,
            "last_used_after": last_after,
            "verified": verified,
            "undo_hint": _UNDO_HINT,
        }
    )
    return result


def _positive_int(value: Any, parameter: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 1:
        raise invalid_request(
            f"Parameter {parameter!r} must be a positive row number.",
            {"parameter": parameter},
        )
    return value


def _worksheet_by_name(target: ResolvedTarget, sheet_name: str | None) -> Any:
    if sheet_name is None:
        try:
            return target.window.ActiveSheet
        except pywintypes.com_error as exc:
            raise _range_error(exc) from exc
    workbook = target.document
    try:
        for index in range(1, int(workbook.Worksheets.Count) + 1):
            candidate = workbook.Worksheets(index)
            if str(candidate.Name).lower() == sheet_name.lower():
                return candidate
    except pywintypes.com_error as exc:
        raise _range_error(exc) from exc
    raise invalid_request(f"Worksheet {sheet_name!r} was not found.", {"sheet": sheet_name})


def _row_values(sheet: Any, row: int, last_col: int | None) -> list[Any]:
    columns = last_col if last_col and last_col >= 1 else 1
    address = f"A{row}:{_column_letter(columns)}{row}" if columns > 1 else f"A{row}"
    try:
        range_object = sheet.Range(address)
        matrix = _rectangular(_matrix(range_object.Value2, 1, columns), 1, columns)
    except pywintypes.com_error as exc:
        raise _range_error(exc) from exc
    return matrix[0]


def move_rows(
    target: ResolvedTarget,
    *,
    row: Any,
    to: Any,
    sheet_name: Any,
    dry_run: bool,
    params: dict[str, Any],
) -> dict[str, Any]:
    """Move one whole row with the choreography proven on this build.

    Rows(src).Copy() puts the row on the clipboard; Rows(dst).Insert() pastes
    the clipboard into the inserted row; clearing CutCopyMode drops it; and
    deleting the duplicated original row completes the move. A failed move
    leaves the sheet half-moved, so the completed steps are always reported:
    re-read the sheet before cleaning up, and note that a single undo only
    reverts the last step.
    """
    check_expected_document(target, params)
    row_number = _positive_int(row, "row")
    to_number = _positive_int(to, "to")
    if row_number == to_number:
        raise invalid_request(
            "Parameter 'to' must differ from 'row'.", {"row": row_number, "to": to_number}
        )
    if sheet_name is not None and (not isinstance(sheet_name, str) or not sheet_name.strip()):
        raise invalid_request("Parameter 'sheet' must be a sheet name.", {"parameter": "sheet"})
    sheet = _worksheet_by_name(target, sheet_name)
    last_row = _last_used_index(sheet, "rows")
    last_col = _last_used_index(sheet, "columns")
    if last_row is not None:
        if row_number > last_row:
            raise invalid_request(
                f"Row {row_number} is below the last used row ({last_row}).",
                {"row": row_number, "last_used_row": last_row},
            )
        if to_number > last_row + 1:
            raise invalid_request(
                f"Parameter 'to' is beyond the end of the used area ({last_row} + 1).",
                {"to": to_number, "last_used_row": last_row},
            )
    direction = "up" if to_number < row_number else "down"
    # Inserting above the source shifts it down (delete source + 1); inserting
    # below leaves it in place (delete source). For a downward move the row is
    # inserted one past the destination so it lands exactly at 'to' after the
    # source delete shifts everything up by one.
    insert_at = to_number if to_number < row_number else to_number + 1
    delete_after_insert = row_number + 1 if to_number < row_number else row_number
    if insert_at > 1048576:
        raise invalid_request(
            "The destination row is beyond the end of the worksheet.",
            {"to": to_number},
        )
    result: dict[str, Any] = {
        "sheet": _sheet_name(sheet),
        "row": row_number,
        "to": to_number,
        "direction": direction,
        "steps": ["copy", "insert", "clear_clipboard", "delete_original"],
        "dry_run": dry_run,
        "applied": False,
        "completed_steps": [],
    }
    if dry_run:
        return result
    require_confirm(params, "excel.rows.move")

    completed: list[str] = []
    original_sheet = None
    try:
        if sheet_name is not None:
            active = target.window.ActiveSheet
            if _sheet_name(active).lower() != _sheet_name(sheet).lower():
                original_sheet = active
                sheet.Activate()
                completed.append("activate_sheet")
        values = _row_values(sheet, row_number, last_col)
        sheet.Rows(row_number).Copy()
        completed.append("copy")
        sheet.Rows(insert_at).Insert(_XL_INSERT_SHIFT["down"])
        completed.append("insert")
        target.application.CutCopyMode = False
        completed.append("clear_clipboard")
        sheet.Rows(delete_after_insert).Delete()
        completed.append("delete_original")
    except pywintypes.com_error as exc:
        try:
            target.application.CutCopyMode = False
        except Exception:
            pass
        raise BridgeError(
            ErrorCode.RANGE_INVALID,
            "The row move stopped partway. Re-read the sheet before continuing: "
            "it may be half-moved, and a single undo only reverts the last step.",
            400,
            {"error": str(exc), "completed_steps": completed},
        ) from exc
    finally:
        if original_sheet is not None:
            try:
                original_sheet.Activate()
            except Exception:
                pass

    moved_values = _row_values(sheet, to_number, last_col)
    result.update(
        {
            "applied": True,
            "completed_steps": completed,
            "verified": moved_values == values,
            "undo_note": (
                "the move ran as separate Office actions; one undo reverts only "
                "the last one"
            ),
        }
    )
    return result


# ---------------------------------------------------------------------------
# Read: cell formatting
#
# Multi-cell reads are unreliable on this build (Interior.Color reports 0 for
# mixed ranges) and colors are stored BGR, so every cell is read individually
# and colors come back as RRGGBB hex.
# ---------------------------------------------------------------------------


def _color_hex(value: Any) -> str | None:
    try:
        raw = int(value) & 0xFFFFFF
    except (TypeError, ValueError):
        return None
    return f"{raw & 0xFF:02X}{(raw >> 8) & 0xFF:02X}{(raw >> 16) & 0xFF:02X}"


def _cell_format(cell: Any, include_borders: bool) -> dict[str, Any]:
    entry: dict[str, Any] = {}
    try:
        pattern = int(cell.Interior.Pattern)
    except Exception:
        pattern = _XL_PATTERN_NONE
    if pattern == _XL_PATTERN_NONE:
        entry["fill"] = None
    else:
        try:
            entry["fill"] = _color_hex(cell.Interior.Color)
        except Exception:
            entry["fill"] = None
    font: dict[str, Any] = {}
    try:
        font["name"] = str(cell.Font.Name)
    except Exception:
        font["name"] = None
    try:
        font["size"] = float(cell.Font.Size)
    except Exception:
        font["size"] = None
    try:
        font["bold"] = bool(cell.Font.Bold)
    except Exception:
        font["bold"] = None
    try:
        font["color"] = _color_hex(cell.Font.Color)
    except Exception:
        font["color"] = None
    entry["font"] = font
    try:
        entry["number_format"] = str(cell.NumberFormat)
    except Exception:
        entry["number_format"] = None
    try:
        entry["halign"] = int(cell.HorizontalAlignment)
    except Exception:
        entry["halign"] = None
    try:
        entry["valign"] = int(cell.VerticalAlignment)
    except Exception:
        entry["valign"] = None
    try:
        entry["wrap"] = bool(cell.WrapText)
    except Exception:
        entry["wrap"] = None
    if include_borders:
        borders: dict[str, Any] = {}
        for side, position in _XL_BORDER_SIDES.items():
            try:
                border = cell.Borders(position)
                style = int(border.LineStyle)
            except Exception:
                borders[side] = None
                continue
            item: dict[str, Any] = {"style": _XL_LINE_STYLES.get(style, style)}
            if style != _XL_PATTERN_NONE:
                try:
                    item["weight"] = int(border.Weight)
                except Exception:
                    item["weight"] = None
                try:
                    item["color"] = _color_hex(border.Color)
                except Exception:
                    item["color"] = None
            borders[side] = item
        entry["borders"] = borders
    return entry


def read_format(
    target: ResolvedTarget,
    *,
    address: str,
    max_cells: int,
    include_borders: bool = False,
) -> dict[str, Any]:
    range_object = _range(target, address)
    rows, columns = _dims(range_object)
    check_cell_limit(rows * columns, max_cells)
    try:
        top_row = int(range_object.Row)
        left_column = int(range_object.Columns(1).Column)
    except pywintypes.com_error as exc:
        raise _range_error(exc) from exc
    cells: list[dict[str, Any]] = []
    for row in range(1, rows + 1):
        for column in range(1, columns + 1):
            try:
                cell = range_object.Cells(row, column)
            except pywintypes.com_error as exc:
                raise _range_error(exc) from exc
            entry = {
                "cell": f"{_column_letter(left_column + column - 1)}{top_row + row - 1}"
            }
            entry.update(_cell_format(cell, include_borders))
            cells.append(entry)
    return {
        "address": _address(range_object),
        "rows": rows,
        "columns": columns,
        "cell_count": rows * columns,
        "borders": include_borders,
        "cells": cells,
    }
