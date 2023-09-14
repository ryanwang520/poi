import datetime
import re
from functools import singledispatch
from inspect import signature
from typing import Any, Callable, Dict

from ..nodes import Cell, Col, Image, Row, Table
from ..utils import get_obj_attr
from ..writer import Writer


def call_by_sig(fn: Callable[..., Any], *args: Any) -> Any:
    sig = signature(fn)
    arg_length = len(sig.parameters)
    return fn(*args[:arg_length])


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
        for i, column in enumerate(self.columns):
            width = column.width or self.col_width
            if width:
                writer.worksheet.set_column(self.col + i, self.col + i, width)
            if should_write(column.title):
                writer.write(row, col + i, column.title, self.cell_format)

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
                        writer.write(
                            row + i + 1, col + j, val, {**self.cell_format, **fmt}
                        )

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

    return visitor
