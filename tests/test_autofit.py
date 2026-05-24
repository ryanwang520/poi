import datetime
import os
from pathlib import Path
from typing import NamedTuple, Optional

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


def test_autofit_mixed_widths(pytestconfig):
    class Product(NamedTuple):
        id: int
        name: str
        category: str
        description: str

    data = [
        Product(1, "Phone", "Electronics", "A smartphone with great camera"),
        Product(2, "Book", "Office", "An interesting paperback novel"),
    ]

    columns = [
        # Explicit static width
        {"attr": "id", "title": "ID", "width": 8},
        # Explicit auto width
        {"attr": "name", "title": "Product Name", "width": "auto"},
        # No width, using Table default which we'll set to 25
        {"attr": "category", "title": "Category Name"},
        # No width, using Table default which is "auto" (will be overridden)
        {"attr": "description", "title": "Long Description", "width": "auto"},
    ]

    sheet = Sheet(
        root=Table(
            data=data,
            columns=columns,
            col_width=25,  # Table level default is static 25
            border=1,
        )
    )

    if pytestconfig.getoption("update_snapshot"):
        sheet.write("tests/__snapshots__/autofit_mixed_widths.xlsx")
    else:
        assert_match_snapshot(sheet, "autofit_mixed_widths.xlsx")


def test_autofit_empty_and_nulls(pytestconfig):
    class EmptyItem(NamedTuple):
        name: str
        value: Optional[float]
        notes: Optional[str]

    # Test with actual null/None/empty values
    data = [
        EmptyItem(name="", value=None, notes=None),
        EmptyItem(name="Item A", value=12.34, notes="Some notes here"),
        EmptyItem(name="电脑", value=None, notes=""),
    ]

    columns = [
        {"attr": "name", "title": "Name Column", "width": "auto"},
        {"attr": "value", "title": "Value Column", "width": "auto"},
        {"attr": "notes", "title": "Notes Column", "width": "auto"},
    ]

    sheet = Sheet(
        root=Table(
            data=data,
            columns=columns,
            col_width="auto",
            border=1,
        )
    )

    if pytestconfig.getoption("update_snapshot"):
        sheet.write("tests/__snapshots__/autofit_empty_and_nulls.xlsx")
    else:
        assert_match_snapshot(sheet, "autofit_empty_and_nulls.xlsx")


def test_autofit_empty_data(pytestconfig):
    class Item(NamedTuple):
        name: str
        age: int

    # No rows at all
    data = []

    columns = [
        {"attr": "name", "title": "User Name Header", "width": "auto"},
        {"attr": "age", "title": "Age", "width": "auto"},
    ]

    sheet = Sheet(
        root=Table(
            data=data,
            columns=columns,
            col_width="auto",
            border=1,
        )
    )

    if pytestconfig.getoption("update_snapshot"):
        sheet.write("tests/__snapshots__/autofit_empty_data.xlsx")
    else:
        assert_match_snapshot(sheet, "autofit_empty_data.xlsx")


def test_autofit_various_types(pytestconfig):
    class Various(NamedTuple):
        s_eng: str
        s_cjk: str
        val_int: int
        val_float: float
        val_bool: bool
        dt: datetime.datetime
        d: datetime.date
        t: datetime.time

    data = [
        Various(
            s_eng="Short",
            s_cjk="短",
            val_int=1,
            val_float=1.1,
            val_bool=True,
            dt=datetime.datetime(2026, 5, 24, 8, 30, 0),
            d=datetime.date(2026, 5, 24),
            t=datetime.time(8, 30, 0),
        ),
        Various(
            s_eng="This is a very long english string",
            s_cjk="这是一段非常长的中文测试文本，用于测试自适应宽度",
            val_int=9999999,
            val_float=1234567.89,
            val_bool=False,
            dt=datetime.datetime(2026, 12, 31, 23, 59, 59),
            d=datetime.date(2026, 12, 31),
            t=datetime.time(23, 59, 59),
        ),
    ]

    columns = [
        {"attr": "s_eng", "title": "English", "width": "auto"},
        {"attr": "s_cjk", "title": "Chinese", "width": "auto"},
        {"attr": "val_int", "title": "Integer", "width": "auto"},
        {
            "attr": "val_float",
            "title": "Float",
            "width": "auto",
            "format": {"num_format": "$#,##0.00"},
        },
        {"attr": "val_bool", "title": "Boolean", "width": "auto"},
        {"attr": "dt", "title": "Datetime", "width": "auto"},
        {
            "attr": "d",
            "title": "Date",
            "width": "auto",
            "format": {"num_format": "yyyy/m/d"},
        },
        {"attr": "t", "title": "Time", "width": "auto"},
    ]

    sheet = Sheet(
        root=Table(
            data=data,
            columns=columns,
            col_width="auto",
            border=1,
        )
    )

    if pytestconfig.getoption("update_snapshot"):
        sheet.write("tests/__snapshots__/autofit_various_types.xlsx")
    else:
        assert_match_snapshot(sheet, "autofit_various_types.xlsx")


def test_autofit_header_comments(pytestconfig):
    class CommentItem(NamedTuple):
        name: str
        value: int

    data = [
        CommentItem(name="A", value=100),
        CommentItem(name="B", value=200000),
    ]

    columns = [
        {
            "attr": "name",
            "title": "Name",
            "width": "auto",
            "title_comment": "This is a comment for name column",
            "title_comment_options": {"author": "Test Author", "visible": True},
        },
        {
            "attr": "value",
            "title": "Value",
            "width": "auto",
            "title_comment": "Value comment",
        },
    ]

    sheet = Sheet(
        root=Table(
            data=data,
            columns=columns,
            col_width="auto",
            border=1,
        )
    )

    if pytestconfig.getoption("update_snapshot"):
        sheet.write("tests/__snapshots__/autofit_header_comments.xlsx")
    else:
        assert_match_snapshot(sheet, "autofit_header_comments.xlsx")
