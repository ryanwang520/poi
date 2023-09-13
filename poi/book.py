from __future__ import annotations

from io import BytesIO

from .sheet import Sheet
from .visitors.writer import writer_visitor
from .writer import BytesIOWorkBook, Writer


class Book:
    def __init__(self) -> None:
        self.sheets: list[Sheet] = []

    def add_sheet(self, worksheet: Sheet) -> None:
        self.sheets.append(worksheet)

    def write(self, filename: str) -> None:
        data = self.write_to_bytes_io()
        with open(filename, "wb") as f:
            f.write(data.read())
            data.close()

    def write_to_bytes_io(self) -> BytesIO:
        workbook = BytesIOWorkBook()
        for sheet in self.sheets:
            worksheet = workbook.add_worksheet()
            writer = Writer(workbook, worksheet)
            visitor = writer_visitor(writer)
            sheet.root.accept(visitor)

        workbook.close()
        return workbook.io
