# Poi: The Declarative Way to Excel at Excel in Python

![CI](https://github.com/ryanwang520/poi/actions/workflows/tests.yaml/badge.svg)

**Poi** is a declarative, node-based layout engine for Python that makes generating complex, beautifully designed Excel spreadsheets a breeze. Instead of calculating coordinates, cells, rows, and spans manually, you define your spreadsheet as a nesting-friendly tree of components—just like React or HTML—and let the layout engine handle the math.

### ❓ Why the name "Poi"?
In Java, [Apache POI](https://poi.apache.org/) is the legendary standard library for working with Microsoft Office files. Its name originally stood for *"Poor Obfuscation Implementation"* (a tongue-in-cheek reference to Microsoft's binary file formats). 

**Poi for Python** inherits that spirit—making Excel generation dead simple—but elevates it with a modern, declarative, CSS-like component layout system that has never existed in the Python ecosystem before.

### ⚡ When is Poi Indispensable?
You should use **Poi** if you are building:
* 📊 **Complex Dashboard Reports**: Reports with non-tabular headers, KPI summaries, and detailed tables side-by-side or stacked vertically on a single sheet.
* 🎨 **Designer-grade Spreadsheets**: Spreadsheets requiring pixel-perfect column auto-fitting, elegant borders, solid color fills, and conditional highlights.
* 💬 **Interactive Reports**: Sheets that require custom cell/header tooltips (comments) to explain calculations to business users.
* 📂 **Multi-sheet Books**: Grouping disparate dashboards and sales data tables across multiple worksheet tabs seamlessly.

---

### 📊 How Poi Compares (vs xlcompose / xltpl / pandas)

| Feature / Dimension | 🎨 **Poi** | 🐼 **pandas** | 📄 **xltpl** | 📐 **xlcompose** |
| :--- | :--- | :--- | :--- | :--- |
| **Core Philosophy** | **Declarative Node-Tree (React-like)** | Flat Dataframe Export | Excel Template-based | Layout Engine |
| **Complex Layouts** (nested containers, KPI cards + tables) | **Perfect (Automatic `Row`/`Col` nesting)** | Very Hard (Manual cells/offsets) | Hard (Rigid template structure) | Good (Flow-based layouts) |
| **Coordinate Math** (rows/cols index, rowspan/colspan) | **100% Automated** | Manual / Imperative | Rigid / Template-bound | Semi-automated |
| **Column Auto-fit** (CJK double-width & padding support) | **Smart & Native (`col_width="auto"`)** | Manual (`set_column`) | Rigid / Template-bound | Lacking / Inaccurate |
| **Styling & Highlights** | **CSS-like & Dynamic Lambdas** | Verbose Styler API | Pre-styled in Template | Object-oriented |
| **Cell / Header Comments** | **Interactive Tooltips Support** | No | No | No |
| **Multi-sheet Books** | **Yes (`Book` API)** | Yes (`ExcelWriter`) | Yes | Yes |
| **Maintained Status** | **Active & Modern** | Active (Not layout-focused) | Legacy / Template-only | Inactive |

---




## Installation

```bash
pip install poi
```

## Quick start

### Create a sheet object and write to a file.

```python
from poi import Sheet, Cell
sheet = Sheet(
    root=Cell("hello world")
)

sheet.write('hello.xlsx')
```

![hello](https://github.com/baoshishu/poi/raw/master/docs/assets/hello.png)

See, it's pretty simple and clear.


### Create a Dynamic Table with Column Auto-fit & Formatting

```python
from typing import NamedTuple
from datetime import datetime
import random

from poi import Sheet, Table


class Product(NamedTuple):
    name: str
    desc: str
    price: int
    created_at: datetime
    img: str


data = [
    Product(
        name=f"prod {i}",
        desc=f"desc {i}",
        price=random.randint(1, 100),
        created_at=datetime.now(),
        img="./docs/assets/product.jpg",
    )
    for i in range(5)
]
columns = [
    {
        "type": "image",
        "attr": "img",
        "title": "Product Image",
        "options": {"x_scale": 0.27, "y_scale": 0.25},
    },
    ("name", "Name"),
    ("desc", "Description"),
    ("price", "Price"),
    ("created_at", "Create Time"),
]
sheet = Sheet(
    root=Table(
        data=data,
        columns=columns,
        col_width="auto",  # Automatically sizes columns to fit contents perfectly!
        row_height=80,
        cell_style={
            # Apply bold red font if price is over 50
            "font_color: red; bold: true": lambda record, col: col.attr == "price" and record.price > 50
        },
        date_format="yyyy-mm-dd",
        align="center",
        border=1,
    )
)
sheet.write("table.xlsx")
```

![table](https://github.com/baoshishu/poi/raw/master/docs/assets/table.png)

See how simple it is to create complex tables? You just wrote a dynamically sized, auto-fitted Excel table with conditional formatting in just a few lines of code!


### Features

- 🎨 **Declarative DSL**: Define spreadsheets with a simple, intuitive, nesting-friendly node tree.
- 🚀 **Automatic Layout**: Standard layout primitives (`Row`, `Col`, `Cell`, `Image`) automatically compute rowspans, colspans, offsets, and row/column indexes for you.
- 📏 **Smart Column Auto-fit**: Use `col_width="auto"` to automatically size columns based on the longest value, with native support for Chinese/CJK double-width characters and date patterns.
- 💬 **Interactive Comments**: Attach rich pop-up tooltips to specific cells and table headers.
- 📚 **Multi-sheet Workbooks**: Organize multiple sheets together dynamically using the `Book` API.
- 🔥 **High Performance**: Native speed powered by the `xlsxwriter` engine, with a `fast=True` option to skip writing styles for empty cells.


### Documentation

For complete APIs, references, and examples, visit our comprehensive [Poi Documentation Portal](https://ryanwang520.github.io/poi/):

- 📖 **[Introduction & Philosophy](https://ryanwang520.github.io/poi/)**
- 🧩 **[Components Reference & Auto-fit Guide](https://ryanwang520.github.io/poi/components/)**
- 🎨 **[Styling, Borders, & Comments Guide](https://ryanwang520.github.io/poi/styling/)**
- 📚 **[Multi-sheet Workbooks Guide](https://ryanwang520.github.io/poi/multi-sheet/)**
