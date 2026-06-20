import datetime
import re
import unicodedata
from functools import singledispatch
from inspect import signature
from typing import Any, Callable, Dict, List, Optional, Union
from weakref import WeakKeyDictionary

from ..nodes import Cell, Col, Image, Row, Table
from ..utils import get_obj_attr
from ..writer import Writer

_SIG_PARAM_COUNT: "WeakKeyDictionary[Callable[..., Any], int]" = WeakKeyDictionary()


def _param_count(fn: Callable[..., Any]) -> int:
    """Return the number of parameters of ``fn``, caching the (expensive)
    ``inspect.signature`` introspection per function object.

    The cache uses weak references so that entries disappear once the callable
    is garbage-collected, avoiding unbounded growth in long-running processes
    that build fresh lambdas per request.
    """
    try:
        n = _SIG_PARAM_COUNT.get(fn)
    except TypeError:
        # Not weak-referenceable (e.g. some builtins); skip caching.
        return len(signature(fn).parameters)
    if n is None:
        n = len(signature(fn).parameters)
        _SIG_PARAM_COUNT[fn] = n
    return n


def call_by_sig(fn: Callable[..., Any], *args: Any) -> Any:
    return fn(*args[: _param_count(fn)])


def format_excel_date(dt: datetime.datetime, excel_fmt: str) -> str:
    # Render dt through an Excel-style date/time format.  Single-char tokens
    # (d/m/h/s) render without leading zeros to match Excel; doubled tokens
    # (dd/mm/hh/ss) zero-pad.
    tokens_rx = re.compile(
        r"(yyyy|yy|dddd|ddd|dd|d|mmmm|mmm|mm|m|hh|h|ss|s|[aA]/[pP]|[aA][mM]/[pP][mM])",
        re.IGNORECASE,
    )
    matches = list(tokens_rx.finditer(excel_fmt))
    has_am_pm = any(m.group(0).lower() in ("am/pm", "a/p") for m in matches)
    hour = (dt.hour % 12 or 12) if has_am_pm else dt.hour
    date_tokens = ("yyyy", "yy", "dddd", "ddd", "dd", "d", "mmmm", "mmm")

    def render(token: str, idx: int) -> str:
        t = token.lower()
        if t == "yyyy":
            return f"{dt.year:04d}"
        if t == "yy":
            return f"{dt.year % 100:02d}"
        if t == "dddd":
            return dt.strftime("%A")
        if t == "ddd":
            return dt.strftime("%a")
        if t == "dd":
            return f"{dt.day:02d}"
        if t == "d":
            return str(dt.day)
        if t == "mmmm":
            return dt.strftime("%B")
        if t == "mmm":
            return dt.strftime("%b")
        if t == "hh":
            return f"{hour:02d}"
        if t == "h":
            return str(hour)
        if t == "ss":
            return f"{dt.second:02d}"
        if t == "s":
            return str(dt.second)
        if t in ("am/pm", "a/p"):
            return dt.strftime("%p")
        # mm / m — month vs minute disambiguated by neighbours
        is_minute = False
        for p_idx in range(idx - 1, -1, -1):
            p_token = matches[p_idx].group(0).lower()
            if p_token in ("hh", "h"):
                is_minute = True
                break
            if p_token in date_tokens:
                break
        if not is_minute:
            for s_idx in range(idx + 1, len(matches)):
                s_token = matches[s_idx].group(0).lower()
                if s_token in ("ss", "s"):
                    is_minute = True
                    break
                if s_token in date_tokens:
                    break
        value = dt.minute if is_minute else dt.month
        return f"{value:02d}" if t == "mm" else str(value)

    parts = []
    last_idx = 0
    for idx, match in enumerate(matches):
        parts.append(excel_fmt[last_idx : match.start()])
        parts.append(render(match.group(0), idx))
        last_idx = match.end()
    parts.append(excel_fmt[last_idx:])

    return "".join(parts)


def format_excel_number(val: Union[float, int], num_format: str) -> str:
    if not isinstance(val, (int, float)):
        return str(val)

    orig_val = val
    sections = num_format.split(";")

    if orig_val > 0:
        fmt_sec = sections[0]
    elif orig_val < 0:
        if len(sections) > 1:
            fmt_sec = sections[1]
            val = abs(orig_val)
        else:
            fmt_sec = sections[0]
            val = orig_val
    else:  # val == 0
        if len(sections) > 2:
            fmt_sec = sections[2]
        elif len(sections) > 1:
            fmt_sec = sections[0]
        else:
            fmt_sec = sections[0]

    fmt_clean = re.sub(r"\[[a-zA-Z0-9]+\]", "", fmt_sec)
    fmt_clean = re.sub(r"_[a-zA-Z0-9_()]", " ", fmt_clean)
    fmt_clean = re.sub(r"\*[a-zA-Z0-9_ -]", "", fmt_clean)
    fmt_clean = fmt_clean.replace('"', "")

    is_percent = "%" in fmt_clean
    if is_percent:
        val = val * 100

    decimal_match = re.search(r"\.([0#?]+)", fmt_clean)
    decimal_places = len(decimal_match.group(1)) if decimal_match else 0

    has_thousands = "," in re.split(r"\.", fmt_clean)[0]

    spec = ""
    if has_thousands:
        spec += ","
    spec += f".{decimal_places}f"

    formatted_num = f"{val:{spec}}"

    num_pattern = re.search(r"([#0?,.\s]+)", fmt_clean)
    if num_pattern:
        prefix = fmt_clean[: num_pattern.start()]
        suffix = fmt_clean[num_pattern.end() :]
    else:
        prefix = ""
        suffix = ""

    res = f"{prefix}{formatted_num}{suffix}"

    if orig_val < 0 and len(sections) == 1 and "-" not in res and "(" not in res:
        res = "-" + res

    return res


def get_string_width(val: Any, num_format: Optional[str] = None) -> int:
    if val is None:
        return 0
    if isinstance(val, (datetime.datetime, datetime.date, datetime.time)):
        if not num_format:
            if isinstance(val, datetime.datetime):
                return 19
            if isinstance(val, datetime.date):
                return 10
            if isinstance(val, datetime.time):
                return 8
            s = str(val)
        else:
            try:
                if isinstance(val, datetime.datetime):
                    dt = val
                elif isinstance(val, datetime.date):
                    dt = datetime.datetime.combine(val, datetime.time.min)
                else:
                    dt = datetime.datetime.combine(datetime.date(2026, 1, 1), val)
                s = format_excel_date(dt, num_format)
            except Exception:
                if isinstance(val, datetime.datetime):
                    return 19
                if isinstance(val, datetime.date):
                    return 10
                if isinstance(val, datetime.time):
                    return 8
                s = str(val)
    elif isinstance(val, (int, float)) and num_format:
        try:
            s = format_excel_number(val, num_format)
        except Exception:
            s = str(val)
    else:
        s = str(val)
    # Fast path: ASCII text has no fullwidth / wide characters, so its display
    # width equals its length.  This avoids a per-character unicodedata lookup
    # for the overwhelmingly common case.
    if s.isascii():
        return len(s)
    # Count fullwidth / wide (CJK) characters as 2, others as 1.
    return sum(2 if unicodedata.east_asian_width(c) in ("F", "W") else 1 for c in s)


_STYLE_SPLIT_RX = re.compile(r"\s*:\s*")
_STYLE_INT_RX = re.compile(r"^\d+$")
_STYLE_FLOAT_RX = re.compile(r"^\d+\.\d+$")


def format_from_style(style_css: str) -> Dict[str, Any]:
    rv: Dict[str, Any] = {}
    for style in style_css.split(";"):
        style = style.strip()
        if style == "":
            continue
        k, v = _STYLE_SPLIT_RX.split(style)
        vl = v.lower()
        if vl == "true":
            rv[k] = True
        elif vl == "false":
            rv[k] = False
        elif _STYLE_INT_RX.match(v):
            rv[k] = int(v)
        elif _STYLE_FLOAT_RX.match(v):
            rv[k] = float(v)
        else:
            rv[k] = v
    return rv


def writer_visitor(writer: Writer, fast: bool = False) -> Any:
    EMPTY_VALUES = (None, "")

    if fast:

        def should_write(value: object) -> bool:
            return value not in EMPTY_VALUES
    else:

        def should_write(value: object) -> bool:
            return True

    @singledispatch
    def visitor(_: Any) -> None:
        pass

    @visitor.register
    def _(self: Row) -> None:
        for child in self.children:
            visitor(child)

    @visitor.register
    def _(self: Col) -> None:
        for child in self.children:
            visitor(child)

    @visitor.register
    def _(self: Table) -> None:  # type: ignore
        row, col = self.row, self.col
        data = self.data
        columns = self.columns
        n_cols = len(columns)
        cell_format = self.cell_format
        worksheet = writer.worksheet
        write = writer.write
        insert_image = writer.insert_image

        row_height = self.row_height
        if row_height:
            if isinstance(row_height, int):
                for i in range(1, len(data) + 1):
                    worksheet.set_row(row + i, row_height)
            else:
                for i in range(1, len(data) + 1):
                    height = call_by_sig(row_height, data[i - 1], i - 1)  # type: ignore
                    if height:
                        worksheet.set_row(row + i, height)

        column_widths: List[Optional[int]] = []
        for i, column in enumerate(columns):
            width = column.width or self.col_width
            if width == "auto":
                column_widths.append(get_string_width(column.title))
            else:
                column_widths.append(None)
                if width:
                    worksheet.set_column(col + i, col + i, width)

        for i, column in enumerate(columns):
            if should_write(column.title):
                write(row, col + i, column.title, cell_format)
                # Write comment for header if present
                if column.title_comment:
                    comment_opts = column.title_comment_options or {}
                    worksheet.write_comment(
                        row, col + i, column.title_comment, comment_opts
                    )

        # Pre-parse the cell styles once instead of for every cell.
        cell_style = self.cell_style
        static_style_fmt: Optional[Dict[str, Any]] = None
        conditional_styles: List[Any] = []
        if isinstance(cell_style, str):
            static_style_fmt = format_from_style(cell_style)
        else:
            conditional_styles = [
                (condition, format_from_style(styles))
                for styles, condition in cell_style.items()
            ]

        # Hoist per-column attributes out of the row loop.
        col_attrs = [c.attr for c in columns]
        col_renders = [c.render for c in columns]
        col_types = [c.type for c in columns]
        col_options = [c.options for c in columns]
        col_formats = [c.format for c in columns]

        datetime_fmt = self.datetime_format or "yyyy-mm-dd hh:mm:ss"
        date_fmt = self.date_format or "yyyy-mm-dd"
        time_fmt = self.time_format or "hh:mm:ss"

        for i, item in enumerate(data):
            target_row = row + i + 1
            for j in range(n_cols):
                attr = col_attrs[j]
                if attr:
                    val = get_obj_attr(item, attr)
                else:
                    render = col_renders[j]
                    assert render
                    val = call_by_sig(render, item, columns[j])

                if col_types[j] == "image":
                    insert_image(target_row, col + j, val, col_options[j])
                    continue

                if not should_write(val):
                    continue

                # Build the per-cell format with precedence (low -> high):
                # cell_format < cell_style < datetime < column.format
                merged_fmt: Dict[str, Any] = dict(cell_format)
                if static_style_fmt is not None:
                    merged_fmt.update(static_style_fmt)
                elif conditional_styles:
                    column = columns[j]
                    for condition, parsed in conditional_styles:
                        if call_by_sig(condition, item, column):
                            merged_fmt.update(parsed)

                if isinstance(val, datetime.datetime):
                    merged_fmt["num_format"] = datetime_fmt
                if isinstance(val, datetime.date):
                    merged_fmt["num_format"] = date_fmt
                if isinstance(val, datetime.time):
                    merged_fmt["num_format"] = time_fmt

                col_fmt = col_formats[j]
                if col_fmt:
                    merged_fmt.update(col_fmt)

                current_width = column_widths[j]
                if current_width is not None:
                    val_width = get_string_width(val, merged_fmt.get("num_format"))
                    if val_width > current_width:
                        column_widths[j] = val_width

                write(target_row, col + j, val, merged_fmt)

        for j, auto_w in enumerate(column_widths):
            if auto_w is not None:
                final_width = max(auto_w + 3, 10)
                worksheet.set_column(col + j, col + j, final_width)

    @visitor.register
    def _(self: Image) -> None:
        writer.insert_image(self.row, self.col, self.filename, self.options)

    @visitor.register
    def _(self: Cell) -> None:
        colspan = self.colspan or 1
        rowspan = self.rowspan or 1
        if self.width:
            writer.worksheet.set_column(self.col, self.col + colspan - 1, self.width)
        if rowspan == 1 and self.height:
            writer.worksheet.set_row(self.row, self.height)
        if not should_write(self.value):
            return
        if colspan == 1 and rowspan == 1:
            writer.write(self.row, self.col, self.value, self.cell_format)
        else:
            writer.merge_range(
                self.row,
                self.col,
                self.row + rowspan - 1,
                self.col + colspan - 1,
                self.value,
                self.cell_format,
            )

        # Write comment if present
        if self.comment:
            writer.worksheet.write_comment(
                self.row, self.col, self.comment, self.comment_options
            )

    return visitor
