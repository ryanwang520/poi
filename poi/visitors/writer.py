from functools import singledispatch
import datetime
import re

from poi.nodes import Row, Col, Table, Cell


def writer_visitor(writer):
    @singledispatch
    def visitor(_):
        pass

    @visitor.register
    def _(self: Row):
        for child in self.children:
            visitor(child)

    @visitor.register
    def _(self: Col):
        for child in self.children:
            visitor(child)

    @visitor.register
    def _(self: Table):
        def get_obj_attr(obj, field):
            for key in field.split("."):
                obj = getattr(obj, key)
                if obj is None:
                    break
            return obj

        row, col = self.row, self.col

        if self.cell_width:
            writer.worksheet.set_column(
                self.col, self.col + len(self.headers), self.cell_width
            )
        for i, header in enumerate(self.headers):
            writer.write(row, col + i, header[1], self.cell_format)

        for i, item in enumerate(self.data):

            for j, (attr_name, display_name) in enumerate(self.headers):
                fmt = {}
                for style, condition in self.cell_style.items():
                    if condition(item, attr_name):
                        k, v = re.split(r"\s*:\s*", style)
                        fmt[k] = v
                val = get_obj_attr(item, attr_name)
                if self.date_format and isinstance(
                    val, (datetime.date, datetime.datetime)
                ):
                    fmt["num_format"] = self.date_format
                writer.write(row + i + 1, col + j, val, {**self.cell_format, **fmt})

    @visitor.register
    def _(self: Cell):
        colspan = self.colspan or 1
        rowspan = self.rowspan or 1
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
