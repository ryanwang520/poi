from pathlib import Path
import os

from poi import __version__, Sheet, Col, Row, Cell


def test_version():
    assert __version__ == '0.1.0'

def assert_match_snapshot(sheet:Sheet, snapshot):
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
                ),
            ],
        )
    )
    assert_match_snapshot(sheet,  "basic.xlsx")

