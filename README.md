# Poi: Make creating Excel XLSX files fun again.

![travis](https://travis-ci.org/baoshishu/poi.svg?branch=master)

Poi helps you write Excel sheet in a declarative way, ensuring you have a better Excel writing experience.

It only supports Python 3.7+.

[Documentation](https://ryanwang520.github.io/poi/)

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


### Sample for rendering a simple table.

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
        row_height=80,
        cell_style={
            "color: red": lambda record, col: col.attr == "price" and record.price > 50
        },
        date_format="yyyy-mm-dd",
        align="center",
        border=1,
    )
)
sheet.write("table.xlsx")
```

![table](https://github.com/baoshishu/poi/raw/master/docs/assets/table.png)
