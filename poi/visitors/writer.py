import datetime
import re
import unicodedata
from functools import singledispatch
from inspect import signature
from typing import Any, Callable, Dict, List, Optional, Union

from ..nodes import Cell, Col, Image, Row, Table
from ..utils import get_obj_attr
from ..writer import Writer


def call_by_sig(fn: Callable[..., Any], *args: Any) -> Any:
    sig = signature(fn)
    arg_length = len(sig.parameters)
    return fn(*args[:arg_length])


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
    # Count fullwidth / wide (CJK) characters as 2, others as 1.
    return sum(2 if unicodedata.east_asian_width(c) in ("F", "W") else 1 for c in s)


def writer_visitor(writer: Writer, fast: bool = False) -> Any:
    EMPTY_VALUES = (None, "")

    def should_write(value: object) -> bool:
        if not fast:
            return True
        return value not in EMPTY_VALUES

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
        for i in range(len(self.data) + 1):
            height = None
            if self.row_height and i >= 1:
                if isinstance(self.row_height, int):
                    height = self.row_height
                else:
                    height = call_by_sig(self.row_height, self.data[i - 1], i - 1)  # type: ignore
            if height:
                writer.worksheet.set_row(self.row + i, height)

        column_widths: List[Optional[int]] = []
        for i, column in enumerate(self.columns):
            width = column.width or self.col_width
            if width == "auto":
                column_widths.append(get_string_width(column.title))
            else:
                column_widths.append(None)
                if width:
                    writer.worksheet.set_column(self.col + i, self.col + i, width)

        for i, column in enumerate(self.columns):
            if should_write(column.title):
                writer.write(row, col + i, column.title, self.cell_format)
                # Write comment for header if present
                if column.title_comment:
                    comment_opts = column.title_comment_options or {}
                    writer.worksheet.write_comment(
                        row, col + i, column.title_comment, comment_opts
                    )

        for i, item in enumerate(self.data):

            def format_from_style(style_css: str) -> Dict[str, Any]:
                rv = {}
                for style in style_css.split(";"):
                    style = style.strip()
                    if style == "":
                        continue
                    k, v = re.split(r"\s*:\s*", style)
                    rv[k] = v
                return rv

            for j, column in enumerate(self.columns):
                if column.attr:
                    val = get_obj_attr(item, column.attr)
                else:
                    assert column.render
                    val = call_by_sig(column.render, item, column)

                if column.type == "image":
                    writer.insert_image(row + i + 1, col + j, val, column.options)
                else:
                    if should_write(val):
                        fmt = {}
                        if isinstance(self.cell_style, str):
                            fmt.update(format_from_style(self.cell_style))

                        else:
                            for styles, condition in self.cell_style.items():
                                if call_by_sig(condition, item, column):
                                    fmt.update(format_from_style(styles))
                        if isinstance(val, datetime.datetime):
                            fmt["num_format"] = (
                                self.datetime_format or "yyyy-mm-dd hh:mm:ss"
                            )
                        if isinstance(val, datetime.date):
                            fmt["num_format"] = self.date_format or "yyyy-mm-dd"
                        if isinstance(val, datetime.time):
                            fmt["num_format"] = self.time_format or "hh:mm:ss"

                        if column.format:
                            fmt.update(column.format)

                        current_width = column_widths[j]
                        if current_width is not None:
                            val_width = get_string_width(val, fmt.get("num_format"))
                            if val_width > current_width:
                                column_widths[j] = val_width

                        writer.write(
                            row + i + 1, col + j, val, {**self.cell_format, **fmt}
                        )

        for j, auto_w in enumerate(column_widths):
            if auto_w is not None:
                final_width = max(auto_w + 3, 10)
                writer.worksheet.set_column(col + j, col + j, final_width)

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
