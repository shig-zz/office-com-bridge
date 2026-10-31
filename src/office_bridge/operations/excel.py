from __future__ import annotations

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
