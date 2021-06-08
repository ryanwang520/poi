from __future__ import annotations
from io import BytesIO

from xlsxwriter import Workbook
from xlsxwriter.worksheet import Worksheet

from typing import Union, List

from .nodes import Box, BoxInstance, Col
from .visitors.printer import print_visitor
from .visitors.writer import writer_visitor
from .writer import Writer, BytesIOWorkBook


class Sheet:
    def __init__(
        self,
        root: Union[Box, List[Box]],
        start_row: int = 0,
        start_col: int = 0,
        name=None,
    ):
        if isinstance(root, list):
            root = Col(children=root)
        BoxInstance(root, start_row, start_col, None)
        self.root = root
        self.name = name

    @classmethod
    def attach_to_exist_worksheet(
        cls,
        workbook: Workbook,
        worksheet: Worksheet,
        root: Box,
        start_row: int = 0,
        start_col: int = 0,
    ):
        sheet = cls(root, start_row, start_col)
        writer = Writer(workbook, worksheet)
        visitor = writer_visitor(writer)
        sheet.root.accept(visitor)
        return writer

    def write_to_bytes_io(self) -> BytesIO:
        workbook = BytesIOWorkBook()
        worksheet = workbook.add_worksheet()
        writer = Writer(workbook, worksheet)
        visitor = writer_visitor(writer)
        self.root.accept(visitor)
        workbook.close()
        return workbook.io

    # old compat
    def write_to_bytesio(self):
        return self.write_to_bytes_io()

    def write(self, filename: str):
        io = self.write_to_bytes_io()
        with open(filename, "wb") as f:
            f.write(io.read())

    def print(self):
        self.root.accept(print_visitor)

    def to_bytes_io(self):
        return self.write_to_bytes_io()
