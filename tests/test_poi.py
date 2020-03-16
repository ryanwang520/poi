import datetime
from pathlib import Path
import os
from typing import NamedTuple

from poi import __version__, Sheet, Col, Row, Cell, Table
from poi.nodes import Image


def test_version():
    assert __version__ == "0.2.0"


def assert_match_snapshot(sheet: Sheet, snapshot):

    actual = sheet.write_to_bytesio().read()
    path = Path(os.path.dirname(__file__)) / "__snapshots__" / snapshot
    expect = path.read_bytes()
    assert abs(len(expect) - len(actual)) <= 3


def test_basic(pytestconfig):
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
    if pytestconfig.getoption("update_snapshot"):
        sheet.write("tests/__snapshots__/basic.xlsx")
    else:
        assert_match_snapshot(sheet, "basic.xlsx")


def test_table(pytestconfig):
    class Record(NamedTuple):
        name: str
        image: str
        desc: str
        remark: str
        time: datetime.datetime

    data = [
        Record(
            name=f"name {i}",
            desc=f"desc {i}",
            remark=f"remark {i}",
            time=datetime.datetime.now(),
            image="assets/image.jpeg",
        )
        for i in range(3)
    ]
    columns = [
        ("name", "名称"),
        ("desc", "描述"),
        ("remark", "备注"),
        ("time", "时间"),
        {
            "attr": "image",
            "title": "图片",
            "type": "image",
            "options": {"x_scale": 0.25, "y_scale": 0.25},
        },
    ]
    sheet = Sheet(
        root=Table(
            data=data,
            columns=columns,
            col_width=20,
            row_height=lambda: 100,
            cell_style={
                "bg_color: yellow; text_wrap: true;": lambda record, col: col.attr
                == "name"
                and record.name == "name 1"
            },
            date_format="yyyy-mm-dd",
            align="center",
            border=1,
        )
    )
    if pytestconfig.getoption("update_snapshot"):
        sheet.write("tests/__snapshots__/table.xlsx")
    else:
        assert_match_snapshot(sheet, "table.xlsx")


def test_complex_row(pytestconfig):
    sheet = Sheet(
        root=Col(
            # colspan=8,
            children=[
                Row(
                    # offset=1,
                    children=[
                        Cell(
                            "cell 1",
                            offset=2,
                            grow=True,
                            bg_color="yellow",
                            align="center",
                            border=1,
                        )
                    ]
                ),
                Row(
                    children=[
                        Cell("cell 2", bg_color="red", align="top", border=2),
                        Cell(
                            "cell 3",
                            offset=1,
                            grow=True,
                            bg_color="green",
                            align="bottom",
                        ),
                        Col(
                            offset=1,
                            children=[Cell("cell 4", colspan=2), Cell("cell 5")],
                            bg_color="red",
                        ),
                    ]
                ),
                Row(
                    # offset=1,
                    children=[
                        Cell("cell 6", colspan=12, bg_color="cyan", align="center")
                    ]
                ),
            ]
        )
    )
    if pytestconfig.getoption("update_snapshot"):
        sheet.write("tests/__snapshots__/complex_row.xlsx")
    else:
        assert_match_snapshot(sheet, "complex_row.xlsx")


def test_complex_col(pytestconfig):
    sheet = Sheet(
        root=Row(
            # rowspan=8,
            children=[
                Col(
                    # offset=1,
                    children=[
                        Cell(
                            "cell 1",
                            offset=2,
                            grow=True,
                            bg_color="yellow",
                            align="center",
                            border=1,
                        )
                    ]
                ),
                Col(
                    children=[
                        Cell("cell 2", bg_color="red", align="top", border=2),
                        Cell(
                            "cell 3",
                            offset=1,
                            grow=True,
                            bg_color="green",
                            align="bottom",
                        ),
                        Row(
                            offset=1,
                            children=[Cell("cell 4", rowspan=2), Cell("cell 5")],
                            bg_color="red",
                        ),
                    ]
                ),
                Col(
                    # rowspan=8,
                    children=[
                        Cell("cell 6", rowspan=10, bg_color="cyan", align="center")
                    ]
                ),
            ]
        )
    )
    if pytestconfig.getoption("update_snapshot"):
        # sheet.print()
        sheet.write("tests/__snapshots__/complex_col.xlsx")
    else:
        assert_match_snapshot(sheet, "complex_col.xlsx")


def test_image(pytestconfig):
    sheet = Sheet(
        root=Col(
            colspan=8,
            children=[
                Row(
                    rowspan=5,
                    children=[
                        Cell("cell 1", bg_color="cyan", align="center"),
                        Image(
                            "assets/image.jpeg",
                            colspan=4,
                            offset=2,
                            grow=True,
                            align="center",
                            border=1,
                            options={"x_scale": 0.2, "y_scale": 0.2},
                        ),
                        Cell("cell 2", bg_color="red", align="center"),
                    ],
                )
            ],
        )
    )
    if pytestconfig.getoption("update_snapshot"):
        sheet.write("tests/__snapshots__/image.xlsx")
    else:
        assert_match_snapshot(sheet, "image.xlsx")
