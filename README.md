# Poi: Make creating Excel XLSX files fun again.

![travis](https://travis-ci.org/baoshishu/poi.svg?branch=master)

Poi helps you write Excel sheet in a declarative way, ensuring you have a better Excel writing experience.

It only supports Python 3.7+.

## Quick start

Create a sheet object and write to a file.

```python
from poi import Sheet
sheet = Sheet(
    root=Col(
        colspan=8,
        children=[
            Row(
                children=[
                    Cell(
                        "hello",
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
sheet.write('hello.xlsx')
```

See, it's pretty simple and clear.

Sample for rendering a simple table.

```python
class Record(NamedTuple):
    name: str
    desc: str
    remark: str

data = [
    Record(name=f"name {i}", desc=f"desc {i}", remark=f"remark {i}")
    for i in range(3)
]
columns = [("name", "名称"), ("desc", "描述"), ("remark", "备注")]
sheet = Sheet(
    root=Table(
        data=data,
        columns=columns,
        cell_width=20,
        cell_style={
            "bg_color: yellow": lambda record, col: col.attr == "name"
            and record.name == "name 1"
        },
        date_format="yyyy-mm-dd",
        align="center",
        border=1,
    )
)
sheet.write('table.xlsx')
```



