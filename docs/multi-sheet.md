# Multi-sheet Workbooks

While a `Sheet` represents a single worksheet within an Excel file, many real-world reports require multiple tabs/worksheets. 

Poi provides the `Book` class to easily group multiple `Sheet` objects together and write them into a single `.xlsx` file.

---

## The `Book` Class

The `Book` class serves as a container for your worksheets. 

### Methods

- `add_sheet(sheet: Sheet)`: Registers a worksheet in the workbook.
- `write(filename: str)`: Renders and saves the multi-sheet workbook to a local file.
- `write_to_bytes_io() -> BytesIO`: Renders the workbook in memory and returns a `BytesIO` stream (ideal for web responses or cloud storage).

---

## Complete Multi-sheet Example

The following example demonstrates how to create a corporate spreadsheet containing two worksheets: a **Dashboard summary** and a **Detailed Sales Table**.

```python
from typing import NamedTuple
from datetime import date
from poi import Book, Sheet, Col, Row, Cell, Table

# 1. Define detailed data and columns for the second worksheet
class Sale(NamedTuple):
    product: str
    amount: float
    sale_date: date

sales_data = [
    Sale("Enterprise SaaS", 50000.00, date(2026, 5, 1)),
    Sale("Consulting Package", 15000.00, date(2026, 5, 12)),
    Sale("Hardware Upgrade", 8500.25, date(2026, 5, 20)),
]

sales_columns = [
    ("product", "Product/Service"),
    {"attr": "amount", "title": "Revenue", "format": {"num_format": "$#,##0.00"}},
    {"attr": "sale_date", "title": "Date Completed"}
]

# 2. Build the first sheet: Dashboard Summary
dashboard_sheet = Sheet(
    root=Col(
        children=[
            Row(
                children=[
                    Cell(
                        "Q2 Corporate Performance Dashboard",
                        grow=True,
                        bold=True,
                        font_size=16,
                        font_color="#FFFFFF",
                        bg_color="#1F4E78",
                        align="center",
                        height=40
                    )
                ]
            ),
            Row(
                colspan=4,
                children=[
                    Cell("Total Q2 Revenue:", bold=True, offset=1),
                    Cell(73500.25, num_format="$#,##0.00", bold=True, bg_color="#E2EFDA")
                ],
                height=25
            ),
            Row(
                colspan=4,
                children=[
                    Cell("Quarter Status:", bold=True, offset=1),
                    Cell("ON TRACK", bold=True, font_color="#385723", bg_color="#E2EFDA", align="center")
                ],
                height=25
            )
        ]
    )
)
# Note: In a future release, you will be able to customize sheet names!
# Currently sheets receive default sheet names ("Sheet1", "Sheet2", etc.)

# 3. Build the second sheet: Sales Details
details_sheet = Sheet(
    root=Table(
        data=sales_data,
        columns=sales_columns,
        col_width="auto",  # Enables auto-fit for accurate widths
        border=1
    )
)

# 4. Bind the sheets together into a Book and write to disk
book = Book()
book.add_sheet(dashboard_sheet)
book.add_sheet(details_sheet)

book.write("corporate_report.xlsx")
```

---

## Streaming Multi-sheet Books via Web Frameworks

You can render a `Book` to an in-memory stream using `book.write_to_bytes_io()` and send it as a file attachment download in any standard Python web framework.

### Flask Integration

```python
from flask import send_file

@app.route("/download-report")
def download_report():
    # Build your book...
    book = Book()
    book.add_sheet(dashboard_sheet)
    book.add_sheet(details_sheet)
    
    # Retrieve the stream
    stream = book.write_to_bytes_io()
    
    return send_file(
        stream,
        as_attachment=True,
        download_name="performance_report.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
```

### Django Integration

```python
from django.http import HttpResponse

def download_report(request):
    # Build your book...
    book = Book()
    book.add_sheet(dashboard_sheet)
    book.add_sheet(details_sheet)
    
    # Write to HttpResponse
    response = HttpResponse(
        book.write_to_bytes_io().read(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    response["Content-Disposition"] = "attachment; filename=performance_report.xlsx"
    return response
```
