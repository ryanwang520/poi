# Poi: Make creating Excel XLSX files fun again.

Poi helps you write Excel sheet in a declarative way, ensuring you have a better Excel writing experience.

It only supports 3.7+.

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
                        "first",
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



