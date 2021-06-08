from __future__ import annotations
from io import BytesIO

from xlsxwriter import Workbook
from xlsxwriter.worksheet import Worksheet
import xlsxwriter

from typing import Union, List, Optional

from .nodes import Box, BoxInstance, Col
from .visitors.printer import print_visitor
from .visitors.writer import writer_visitor
from .writer import Writer


class Sheet:
    def __init__(
        self,
        root: Union[Box, List[Box]],
        start_row: int = 0,
        start_col: int = 0,
        name=None,
        workbook: Optional[Workbook] = None,
        worksheet: Optional[Worksheet] = None,
    ):
        if isinstance(root, list):
            root = Col(children=root)
        BoxInstance(root, start_row, start_col, None)
        self.root = root
        self.workbook = workbook
        self.worksheet = worksheet
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
        writer = Writer(workbook, worksheet, attached=True)
        visitor = writer_visitor(writer)
        sheet.root.accept(visitor)
        return writer

    def write_to_bytesio(self) -> Writer:
        writer = Writer(worksheet=self.worksheet, workbook=self.workbook)
        visitor = writer_visitor(writer)
        self.root.accept(visitor)
        writer.close()
        return writer

    def write(self, filename: str):
        writer = self.write_to_bytesio()
        with open(filename, "wb") as f:
            f.write(writer.read())

    def print(self):
        self.root.accept(print_visitor)

    def to_bytes_io(self):
        writer = Writer(worksheet=self.worksheet, workbook=self.workbook)
        visitor = writer_visitor(writer)
        self.root.accept(visitor)
        writer.close()
        return writer.output


class Book:
    def __init__(self):
        self.sheets: List[Sheet] = []

    def add_sheet(self, worksheet: Sheet):
        self.sheets.append(worksheet)

    def write(self, filename: str):
        data = self.to_bytes_io()
        with open(filename, "wb") as f:
            f.write(data.read())
            data.close()

    def to_bytes_io(self):
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output)
        for sheet in self.sheets:
            worksheet = workbook.add_worksheet(name=sheet.name)
            sheet.workbook = workbook
            sheet.worksheet = worksheet
            sheet.write_to_bytesio()
        workbook.close()
        output.seek(0)
        return output
