<div align="center">

# Poi

**Declarative Excel for Python.**
Build complex spreadsheets like you build React components — no coordinates, no `merge_range` math.

[![CI](https://github.com/ryanwang520/poi/actions/workflows/tests.yaml/badge.svg)](https://github.com/ryanwang520/poi/actions/workflows/tests.yaml)
[![PyPI](https://img.shields.io/pypi/v/poi.svg)](https://pypi.org/project/poi/)
[![Python](https://img.shields.io/pypi/pyversions/poi.svg)](https://pypi.org/project/poi/)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Docs](https://img.shields.io/badge/docs-online-brightgreen.svg)](https://ryanwang520.github.io/poi/)

</div>

---

## Before & After

A simple report: a title row spanning four columns, two KPI cards, and a data table below.

**With `xlsxwriter` — you manage every coordinate:**

```python
import xlsxwriter

wb = xlsxwriter.Workbook("report.xlsx")
ws = wb.add_worksheet()

title  = wb.add_format({"bold": True, "font_size": 18, "align": "center"})
kpi    = wb.add_format({"bold": True, "align": "center", "border": 1, "font_size": 14})
header = wb.add_format({"bold": True, "bg_color": "#EFEFEF", "border": 1})
cell   = wb.add_format({"border": 1})

ws.merge_range("A1:D1", "Q4 Report", title)
ws.merge_range("A3:B5", "Revenue $1.28M", kpi)
ws.merge_range("C3:D5", "Growth +18%", kpi)

for col, h in enumerate(["Region", "Q3", "Q4", "Δ"]):
    ws.write(6, col, h, header)

for row, r in enumerate(
    [("North", 320, 410, "+28%"), ("South", 280, 305, "+9%")], start=7
):
    for col, val in enumerate(r):
        ws.write(row, col, val, cell)

wb.close()
```

**With Poi — you describe the structure:**

```python
from poi import Sheet, Col, Row, Cell, Table

regions = [
    {"region": "North", "q3": 320, "q4": 410, "delta": "+28%"},
    {"region": "South", "q3": 280, "q4": 305, "delta": "+9%"},
]

Sheet(root=Col(
    Cell("Q4 Report", colspan=4, style="font_size: 18; bold: true; align: center"),
    Row(
        Cell("Revenue\n$1.28M", colspan=2, style="border: 1; bold: true"),
        Cell("Growth\n+18%",    colspan=2, style="border: 1; bold: true"),
    ),
    Table(
        data=regions,
        columns=[("region", "Region"), ("q3", "Q3"), ("q4", "Q4"), ("delta", "Δ")],
        border=1,
    ),
)).write("report.xlsx")
```

No row indexes. No `merge_range` strings. Insert a new section anywhere — everything below shifts automatically.

---

## Quick Start

```bash
pip install poi
```

```python
from poi import Sheet, Table

data = [
    {"name": "iPhone", "price": 999},
    {"name": "iPad",   "price": 599},
]

Sheet(root=Table(
    data=data,
    columns=[("name", "Product"), ("price", "Price")],
    col_width="auto",
    border=1,
)).write("products.xlsx")
```

---

## When to Use Poi

Poi is built for spreadsheets that aren't just "dump a dataframe":

- **Dashboard reports** with KPI cards stacked above detailed tables
- **Multi-section sheets** where the layout shifts as data changes
- **CJK-heavy layouts** where column widths actually matter
- **Multi-sheet books** with mixed structures across tabs

If `df.to_excel()` is good enough for you, you don't need Poi. If you've ever counted rows by hand or rewritten `merge_range("A3:B5", ...)` because you inserted a row — Poi is for you.

---

## Smart CJK Column Auto-fit

A real differentiator. Most Excel libraries measure column width by character count, which silently breaks for Chinese, Japanese, and Korean text — each CJK character renders at roughly twice the width of an ASCII character.

```python
Table(data=data, columns=[...], col_width="auto")
```

Poi counts double-width characters correctly and recognizes common date patterns. Your `产品名称` column fits on the first try.

---

## Full Example: Images & Conditional Formatting

```python
from typing import NamedTuple
from datetime import datetime
from poi import Sheet, Table


class Product(NamedTuple):
    name: str
    price: int
    created_at: datetime
    img: str


data = [
    Product(f"prod {i}", i * 17, datetime.now(), "./product.jpg")
    for i in range(5)
]

Sheet(root=Table(
    data=data,
    columns=[
        {"type": "image", "attr": "img", "title": "Image",
         "options": {"x_scale": 0.27, "y_scale": 0.25}},
        ("name", "Name"),
        ("price", "Price"),
        ("created_at", "Created"),
    ],
    col_width="auto",
    row_height=80,
    cell_style={
        # Bold red when price > 50
        "font_color: red; bold: true":
            lambda record, col: col.attr == "price" and record.price > 50,
    },
    date_format="yyyy-mm-dd",
    align="center",
    border=1,
)).write("table.xlsx")
```

![table](https://github.com/baoshishu/poi/raw/master/docs/assets/table.png)

---

## Comparison

|                                             | **Poi**                       | `xlsxwriter`                       | `openpyxl`                       | `pandas.to_excel`    | `xltpl`                        |
| ------------------------------------------- | ----------------------------- | ---------------------------------- | -------------------------------- | -------------------- | ------------------------------ |
| **Style**                                   | Declarative tree              | Imperative coordinates             | Imperative coordinates           | One-shot export      | Excel as template              |
| **Nested layouts** (KPI cards above tables) | Native (`Row`/`Col`)          | Manual offset math                 | Manual offset math               | Not supported        | Constrained by template        |
| **Coordinate / span math**                  | Automatic                     | You compute                        | You compute                      | N/A                  | Template-bound                 |
| **CJK-aware column auto-fit**               | Built in (`col_width="auto"`) | Manual `set_column`                | Manual                           | Not supported        | Template-bound                 |
| **Conditional cell style**                  | Lambda per row                | Manual per cell                    | Manual per cell                  | Verbose `Styler` API | Pre-baked in template          |
| **Multi-sheet books**                       | `Book` API                    | Supported                          | Supported                        | `ExcelWriter`        | Supported                      |
| **Best when**                               | Designed reports & dashboards | Maximum control, very large writes | Reading + editing existing files | Pure dataframe dump  | Fixed layouts filled with data |

---

## Features

- 🎨 **Declarative DSL** — describe spreadsheets as a tree of `Row`, `Col`, `Cell`, `Image`, `Table`
- 📐 **Automatic layout** — rowspans, colspans, offsets, and indexes are computed for you
- 📏 **Smart column auto-fit** — `col_width="auto"`, native CJK & date-pattern support
- 🎯 **CSS-like styling** — `style="font_size: 14; bold: true; bg_color: #EEE"`
- 🎛️ **Lambda-based conditional formatting** — style cells by predicate
- 💬 **Interactive comments** — pop-up tooltips on cells and headers
- 📚 **Multi-sheet workbooks** via the `Book` API
- ⚡ **Fast** — built on `xlsxwriter`; `fast=True` skips empty-cell styling for very large outputs

---

## Documentation

📖 **[Read the docs →](https://ryanwang520.github.io/poi/)**

- [Introduction & philosophy](https://ryanwang520.github.io/poi/)
- [Components reference & auto-fit guide](https://ryanwang520.github.io/poi/components/)
- [Styling, borders & comments](https://ryanwang520.github.io/poi/styling/)
- [Multi-sheet workbooks](https://ryanwang520.github.io/poi/multi-sheet/)

---

## Contributing

Poi is a small, focused library and feedback is very welcome — especially from teams using it on real Chinese / multilingual reports. Open an [issue](https://github.com/ryanwang520/poi/issues), share a tricky layout, or send a PR.

---

<sub>_Why "Poi"? A short, easy-to-type name — like a Point on a grid. That's it._</sub>
