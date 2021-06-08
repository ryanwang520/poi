from __future__ import annotations
from io import BytesIO

from xlsxwriter import Workbook
from xlsxwriter.worksheet import Worksheet
import xlsxwriter

from typing import Union, List, IO, Optional

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
        workbook: Optional[Workbook] = None,
        worksheet: Optional[Worksheet] = None,
    ):
        if isinstance(root, list):
            root = Col(children=root)
        BoxInstance(root, start_row, start_col, None)
        self.root = root
        self.workbook = workbook
        self.worksheet = worksheet

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

    def write(self, filename: Union[str, IO]):
        writer = self.write_to_bytesio()
        if isinstance(filename, str):
            with open(filename, "wb") as f:
                f.write(writer.read())
        else:
            filename.write(writer.read())

    def print(self):
        self.root.accept(print_visitor)

    def to_bytes_io(self, output: Optional[IO] = None):
        writer = Writer(worksheet=self.worksheet, workbook=self.workbook, output=output)
        visitor = writer_visitor(writer)
        self.root.accept(visitor)
        writer.close()
        return writer.output


class Book:
    def __init__(self):
        self.sheets: List[Sheet] = []

    def add_worksheet(self, worksheet: Sheet):
        self.sheets.append(worksheet)

    def write(self, filename: str):
        workbook = xlsxwriter.Workbook()
        with open(filename, "wb") as f:
            for sheet in self.sheets:
                worksheet = workbook.add_worksheet()
                sheet.workbook = workbook
                sheet.worksheet = worksheet
                sheet.write(f)

    def to_bytes_io(self):
        workbook = xlsxwriter.Workbook()
        output = BytesIO()
        for sheet in self.sheets:
            worksheet = workbook.add_worksheet()
            sheet.workbook = workbook
            sheet.worksheet = worksheet
            sheet.to_bytes_io(output)
        return output


def main():
    book = Book()
    book.add_worksheet(Sheet())
    book.add_worksheet(Sheet())
    book.add_worksheet(Sheet())
    book.write()
