from __future__ import annotations

from io import BytesIO
from typing import Any

from xlsxwriter import Workbook
from xlsxwriter.worksheet import Worksheet

from .nodes import Box, BoxInstance, Col
from .visitors.printer import print_visitor
from .visitors.writer import writer_visitor
from .writer import BytesIOWorkBook, Writer


class Sheet:
    def __init__(
        self,
        root: Box | list[Box],
        start_row: int = 0,
        start_col: int = 0,
        global_format: dict[str, Any] | None = None,
        fast: bool = False,
    ):
        if isinstance(root, list):
            root = Col(children=root)
        BoxInstance(root, start_row, start_col, None)
        self.root = root
        self.global_format = global_format
        self.fast = fast

    @classmethod
    def attach_to_exist_worksheet(
        cls,
        workbook: Workbook,
        worksheet: Worksheet,
        root: Box,
        start_row: int = 0,
        start_col: int = 0,
        global_format: dict[str, Any] | None = None,
    ) -> Writer:
        sheet = cls(root, start_row, start_col)
        writer = Writer(workbook, worksheet, global_format=global_format)
        visitor = writer_visitor(writer, fast=sheet.fast)
        sheet.root.accept(visitor)
        return writer

    def write_to_bytes_io(self) -> BytesIO:
        workbook = BytesIOWorkBook()
        worksheet = workbook.add_worksheet()
        writer = Writer(workbook, worksheet, self.global_format)
        visitor = writer_visitor(writer, fast=self.fast)
        self.root.accept(visitor)
        workbook.close()
        return workbook.io

    def write_to_worksheet(self, workbook: Workbook, worksheet: Worksheet) -> None:
        writer = Writer(workbook, worksheet, self.global_format)
        visitor = writer_visitor(writer, fast=self.fast)
        self.root.accept(visitor)

    def write(self, filename: str) -> None:
        io = self.write_to_bytes_io()
        with open(filename, "wb") as f:
            f.write(io.read())

    def print(self) -> None:
        self.root.accept(print_visitor)

    def to_bytes_io(self) -> BytesIO:
        return self.write_to_bytes_io()
