from __future__ import annotations

from typing import List

from . import Sheet
from .visitors.writer import writer_visitor
from .writer import Writer, BytesIOWorkBook


class Book:
    def __init__(self):
        self.sheets: List[Sheet] = []

    def add_sheet(self, worksheet: Sheet):
        self.sheets.append(worksheet)

    def write(self, filename: str):
        data = self.write_to_bytes_io()
        with open(filename, "wb") as f:
            f.write(data.read())
            data.close()

    def write_to_bytes_io(self):

        workbook = BytesIOWorkBook()
        for sheet in self.sheets:
            worksheet = workbook.add_worksheet(name=sheet.name)
            writer = Writer(workbook, worksheet)
            visitor = writer_visitor(writer)
            sheet.root.accept(visitor)

        workbook.close()
        return workbook.io
