# Poi: The Declarative Way to Excel at Excel in Python

![CI](https://github.com/ryanwang520/poi/actions/workflows/tests.yaml/badge.svg)

## Why Poi?

Creating Excel files programmatically has always been a chore. Current libraries offer limited flexibility, especially when you need more than a basic table. That's where Poi comes in, offering you a simple, intuitive, yet powerful DSL to make Excel files exactly the way you want them.



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


### Create a Dynamic Table with Conditional Formatting


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

See how simple it is to create complex tables? You just wrote a dynamic Excel table with conditional formatting a few lines of code!


### Features

* 🎉 Declarative: Create Excel files with a simple, intuitive DSL.
* 🔥 Fast: Export large Excel files in seconds.
* 🚀 Flexible Layouts: Create any layout you can imagine with our intuitive Row and Col primitives.


### Documentation

For more details, check our comprehensive [Documentation](https://ryanwang520.github.io/poi/)
