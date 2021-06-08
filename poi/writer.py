from __future__ import annotations
from io import BytesIO
from typing import Protocol, List

import xlsxwriter
from xlsxwriter.format import Format
from xlsxwriter.worksheet import Worksheet


class WorkBook(Protocol):
    def add_format(self, format) -> Format:
        ...

    def worksheets(self) -> List[Worksheet]:
        ...

    def add_worksheet(self, name=None) -> Worksheet:
        ...


class BytesIOWorkBook:
    def __init__(self):
        self.io = BytesIO()
        self.workbook = xlsxwriter.Workbook(self.io)

    def add_format(self, format):
        return self.workbook.add_format(format)

    def close(self):
        self.workbook.close()
        self.io.seek(0)

    def read(self):
        data = self.io.read()
        self.io.close()
        return data

    def worksheets(self):
        return self.workbook.worksheets()

    def add_worksheet(self, name=None):
        return self.workbook.add_worksheet(name)


class Writer:
    def __init__(
        self,
        workbook: WorkBook,
        worksheet: Worksheet,
        global_format=None,
    ):

        self.workbook = workbook
        self.worksheet = worksheet
        self.global_format = self.workbook.add_format(global_format) if global_format else None
        self.global_format_dict = global_format or {}

    def _calc_format(self, cell_format):
        if cell_format is None:
            cell_format = self.global_format
        elif isinstance(cell_format, dict):
            cell_format = self.workbook.add_format(
                {**self.global_format_dict, **cell_format}
            )
        return cell_format

    def _path_args(self, args):
        last_arg = args[-1]
        if isinstance(last_arg, dict):
            args = list(args[:-1]) + [self._calc_format(last_arg)]
        return args

    def write(self, *args):
        args = self._path_args(args)
        self.worksheet.write(*args)

    def merge_range(self, *args):
        args = self._path_args(args)
        self.worksheet.merge_range(*args)

    def insert_image(self, *args):
        self.worksheet.insert_image(*args)
