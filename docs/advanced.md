# Advanced Features

## Use Optional Channing in Column config

```python
from poi import Sheet, Table

data = [
    {'order': {'title': {'name': 'Sample 1'}}},
    {'order': None},
    {'order': {'title': None}},
]

columns = [
    {
        "attr": "order?.title?.name",
        "title": "Order Name",
    },
    # other columns
]

sheet = Sheet(
    root=Table(
        data=data,
        columns=columns,
        # other configurations
    )
)

sheet.write("result.xlsx")
```

With this setup, your Excel sheet will safely display the name of the title of an order, wherever it exists. If any of the nested properties is missing or None, the cell in Excel will remain empty, ensuring that no runtime errors are thrown.

## Use the grow Property for Dynamic Sizing

One of Poi's standout features is its ability to dynamically adjust the size of Row or Col elements within the sheet. By setting the grow property to True, you can instruct a Row or Col to automatically fill the remaining space within its parent container.


### Row Element

Let's assume you want a title row that takes up the full width of the sheet.



```python
from poi import Sheet, Row, Cell

sheet = Sheet(
    root=Row(
        colspan=3,
        children=[
            Cell(""), # empty cell to fill the first column
            Cell("My Awesome Spreadsheet", grow=True) # cell that will fill the remaining 2 columns
        ]
    )
)

sheet.write('dynamic_sizing_row_example.xlsx')
```

### Col Element

Or maybe you want a column that auto-fills all available vertical space.



```python
from poi import Sheet, Col, Cell

sheet = Sheet(
    root=Col(
        rowspan=3,
        children=[
            Cell(""), # empty cell to fill the first row
            Cell("Data", grow=True) # cell that will fill the remaining 2 rows
        ]
    )
)

sheet.write('dynamic_sizing_col_example.xlsx')
```

In both examples, the Row or Col containing the Cell with grow=True will auto-fill to occupy all available space within their parent containers.



## Fast Mode

Generally Poi will write each cell even the content is empty, but if you're for better performance and don't care about the visual(background, border) of empty cells, you can enable fast mode to skip the empty cells.

```python
from poi import Sheet, Table
sheet = Sheet(
    root=Table(
        data=[],
        columns=[],
    ),
    fast=True
)

```