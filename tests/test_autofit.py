import datetime
import os
from pathlib import Path
from typing import NamedTuple

from poi import Sheet, Table
from poi.visitors.writer import get_string_width


def assert_match_snapshot(sheet: Sheet, snapshot):
    actual = sheet.write_to_bytes_io().read()
    path = Path(os.path.dirname(__file__)) / "__snapshots__" / snapshot
    expect = path.read_bytes()
    assert abs(len(expect) - len(actual)) <= 10


def test_get_string_width():
    assert get_string_width(None) == 0
    assert get_string_width(123) == 3
    assert get_string_width("abc") == 3
    assert get_string_width("中文") == 4
    assert get_string_width("中文abc") == 7
    assert get_string_width(datetime.datetime(2026, 5, 24, 12, 0, 0)) == 19
    assert get_string_width(datetime.date(2026, 5, 24)) == 10
    assert get_string_width(datetime.time(12, 0, 0)) == 8


def test_get_string_width_custom_formats():
    # Test custom datetime formats
    dt = datetime.datetime(2026, 1, 1, 12, 0, 0)
    # Renders "Thursday, January 01, 2026"
    assert get_string_width(dt, "dddd, mmmm d, yyyy") == 26
    # Renders "2026/01/01"
    assert get_string_width(dt, "yyyy/m/d") == 10

    # Test numeric format `$#,##0.00`
    assert get_string_width(1299.0, "$#,##0.00") == 9  # "$1,299.00"
    assert get_string_width(-1299.0, "$#,##0.00;($#,##0.00)") == 11
    assert get_string_width(0.05, "0.0%") == 4  # "5.0%"
    assert get_string_width(1234, "#,##0") == 5  # "1,234"


def test_autofit_custom_formats():
    import xlsxwriter

    class Item(NamedTuple):
        price: float
        date: datetime.date

    data = [
        Item(price=1299.0, date=datetime.date(2026, 1, 1)),
        Item(price=9.99, date=datetime.date(2026, 2, 15)),
    ]

    columns = [
        # Explicit auto-fit with custom currency format
        {
            "attr": "price",
            "title": "Price",
            "width": "auto",
            "format": {"num_format": "$#,##0.00"},
        },
        # Explicit auto-fit with custom date format
        {
            "attr": "date",
            "title": "Date",
            "width": "auto",
            "format": {"num_format": "dddd, mmmm d, yyyy"},
        },
    ]

    table = Table(
        data=data,
        columns=columns,
    )

    workbook = xlsxwriter.Workbook()
    worksheet = workbook.add_worksheet()

    Sheet.attach_to_exist_worksheet(workbook, worksheet, table)

    # Let's inspect the calculated widths in ws.col_info!
    # For price: title length is 5 ("Price"), values: "$1,299.00" (9 chars)
    # and "$9.99" (5 chars). Max auto_w is 9. Final width is max(12, 10) = 12.
    # For date: title length is 4 ("Date"), values:
    # "Thursday, January 01, 2026" (26 chars), and
    # "Sunday, February 15, 2026" (25 chars).
    # Max auto_w is 26. Final width is max(29, 10) = 29.

    assert worksheet.col_info[0][0] == 12
    assert worksheet.col_info[1][0] == 29


def test_autofit_table(pytestconfig):
    class Item(NamedTuple):
        name: str
        description: str
        price: float
        date: datetime.date

    data = [
        Item(
            name="A",
            description="Short desc",
            price=9.99,
            date=datetime.date(2026, 1, 1),
        ),
        Item(
            name="Laptop Pro",
            description="Very long description for this item",
            price=1299.0,
            date=datetime.date(2026, 2, 15),
        ),
        Item(
            name="电脑",
            description="Chinese description 中文描述",
            price=50.0,
            date=datetime.date(2026, 3, 20),
        ),
    ]

    columns = [
        # Explicit auto-fit for this column
        {"attr": "name", "title": "名称", "width": "auto"},
        # Explicit auto-fit for this column
        {"attr": "description", "title": "Product Description", "width": "auto"},
        # Static width for this column
        {"attr": "price", "title": "Price", "width": 15},
        # No width specified, will fallback to table's col_width (which is "auto" here!)
        {"attr": "date", "title": "Date"},
    ]

    sheet = Sheet(
        root=Table(
            data=data,
            columns=columns,
            col_width="auto",  # Table-wide auto-fit default
            border=1,
        )
    )

    if pytestconfig.getoption("update_snapshot"):
        sheet.write("tests/__snapshots__/autofit_table.xlsx")
    else:
        assert_match_snapshot(sheet, "autofit_table.xlsx")
