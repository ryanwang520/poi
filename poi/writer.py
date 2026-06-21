from __future__ import annotations

import json
import logging
from io import BytesIO
from typing import Any, Protocol

import xlsxwriter
from xlsxwriter.format import Format
from xlsxwriter.worksheet import Worksheet

logger = logging.getLogger(__name__)

# Excel stores numbers as IEEE 754 doubles, so integers outside the range
# [-(2**53 - 1), 2**53 - 1] cannot be represented exactly and silently lose
# precision (e.g. 222222222222222222222222222222 -> 2.222222222222222E+29).
# Write such integers as text instead to preserve every digit.
MAX_SAFE_INTEGER = 2**53 - 1


def _coerce_large_int(value: Any) -> Any:
    # `type(value) is int` deliberately excludes bool (a subclass of int).
    if type(value) is int and not -MAX_SAFE_INTEGER <= value <= MAX_SAFE_INTEGER:
        return str(value)
    return value


class WorkBook(Protocol):
    def add_format(self, format: dict[str, Any]) -> Format: ...

    def worksheets(self) -> list[Worksheet]: ...

    def add_worksheet(self, name: str | None = None) -> Worksheet: ...


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
            try:
                fmt_key: Any = tuple(sorted(cell_format.items()))
            except TypeError:
                # Unhashable style values — fall back to a stable string key.
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
        out = list(self._path_args(args))
        # out: (row, col, value, [format])
        if len(out) >= 3:
            out[2] = _coerce_large_int(out[2])
        self.worksheet.write(*out)

    def merge_range(self, *args: Any) -> None:
        out = list(self._path_args(args))
        # out: (first_row, first_col, last_row, last_col, value, [format])
        if len(out) >= 5:
            out[4] = _coerce_large_int(out[4])
        self.worksheet.merge_range(*out)

    def insert_image(self, *args: Any) -> None:
        self.worksheet.insert_image(*args)
