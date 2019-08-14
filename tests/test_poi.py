from pathlib import Path
import os
from typing import NamedTuple

from poi import __version__, Sheet, Col, Row, Cell, Table


def test_version():
    assert __version__ == "0.1.13"


def assert_match_snapshot(sheet: Sheet, snapshot):
    actual = sheet.write_to_bytesio().read()
    path = Path(os.path.dirname(__file__)) / "__snapshots__" / snapshot
    expect = path.read_bytes()
    assert abs(len(expect) - len(actual)) <= 1


def test_basic():
    sheet = Sheet(
        root=Col(
            colspan=8,
            children=[
                Row(
                    children=[
                        Cell(
                            "hello world",
                            offset=2,
                            grow=True,
                            bg_color="yellow",
                            align="center",
                            border=1,
                        )
                    ]
                )
            ],
        )
    )
    assert_match_snapshot(sheet, "basic.xlsx")
    # sheet.write("tests/__snapshots__/basic.xlsx")


def test_table():
    class Record(NamedTuple):
        name: str
        desc: str
        remark: str

    data = [
        Record(name=f"name {i}", desc=f"desc {i}", remark=f"remark {i}")
        for i in range(3)
    ]
    columns = [("name", "名称"), ("desc", "描述"), ("remark", "备注")]
    sheet = Sheet(
        root=Table(
            data=data,
            columns=columns,
            cell_width=20,
            cell_style={
                "bg_color: yellow": lambda record, col: col.attr == "name"
                and record.name == "name 1"
            },
            date_format="yyyy-mm-dd",
            align="center",
            border=1,
        )
    )
    assert_match_snapshot(sheet, "table.xlsx")
    # sheet.write("tests/__snapshots__/table.xlsx")
