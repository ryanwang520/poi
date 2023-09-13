from __future__ import annotations

import json
import logging
from io import BytesIO
from typing import Any, Protocol

import xlsxwriter
from xlsxwriter.format import Format
from xlsxwriter.worksheet import Worksheet

logger = logging.getLogger(__name__)


class WorkBook(Protocol):
    def add_format(self, format: dict[str, Any]) -> Format:
        ...

    def worksheets(self) -> list[Worksheet]:
        ...

    def add_worksheet(self, name: str | None = None) -> Worksheet:
        ...


class BytesIOWorkBook:
    def __init__(self) -> None:
        self.io = BytesIO()
        self.workbook = xlsxwriter.Workbook(self.io)

    def add_format(self, format: dict[str, Any]) -> Format:
        return self.workbook.add_format(format)

    def close(self) -> None:
        self.workbook.close()
        self.io.seek(0)

    def read(self) -> bytes:
        data = self.io.read()
        self.io.close()
        return data

    def worksheets(self) -> Worksheet:
        return self.workbook.worksheets()

    def add_worksheet(self, name: str | None = None) -> Worksheet:
        return self.workbook.add_worksheet(name)


class Writer:
    def __init__(
        self,
        workbook: WorkBook,
        worksheet: Worksheet,
        global_format: dict[str, Any] | None = None,
    ) -> None:
        self.workbook = workbook
        self.worksheet = worksheet
        self.global_format = (
            self.workbook.add_format(global_format) if global_format else None
        )
        self.global_format_dict = global_format or {}
        self.formats: dict[str, Any] = {}

    def _calc_format(self, cell_format: Any) -> Any:
        if not cell_format:
            return self.global_format
        elif isinstance(cell_format, dict):
            fmt_key = json.dumps(cell_format, sort_keys=True)
            fmt = self.formats.get(fmt_key)
            if not fmt:
                fmt = self.workbook.add_format(
                    {**self.global_format_dict, **cell_format}
                )
                self.formats[fmt_key] = fmt
            return fmt
        else:
            logger.error(f"cell_format must be dict, got {cell_format}")
            return self.global_format

    def _path_args(self, args: Any) -> Any:
        last_arg = args[-1]
        if isinstance(last_arg, dict):
            args = list(args[:-1]) + [self._calc_format(last_arg)]
        return args

    def write(self, *args: Any) -> None:
        args = self._path_args(args)
        self.worksheet.write(*args)

    def merge_range(self, *args: Any) -> None:
        args = self._path_args(args)
        self.worksheet.merge_range(*args)

    def insert_image(self, *args: Any) -> None:
        self.worksheet.insert_image(*args)
