from io import BytesIO

import xlsxwriter

from . import Sheet


class Writer:
    worksheet: Sheet

    def __init__(
        self,
        workbook=None,
        worksheet=None,
        global_format=None,
        attached=False,
        output=None,
    ):

        self.output = (output or BytesIO()) if not attached else None
        if not (workbook and worksheet):
            self.workbook = xlsxwriter.Workbook(self.output)
            self.worksheet = self.workbook.add_worksheet()
        else:
            self.workbook = workbook
            self.worksheet = worksheet
        self.global_format = self.workbook.add_format(global_format)
        self.global_format_dict = global_format or {}

    def close(self):
        self.workbook.close()
        if self.output:
            self.output.seek(0)

    def read(self):
        if not self.output:
            raise ValueError("cannot read from attached worksheet")
        return self.output.read()

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
