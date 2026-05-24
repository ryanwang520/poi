# Components Reference

Poi provides a declarative layout engine based on a node tree. Every visible element on the sheet is represented by a `Box` subclass.

---

## The Core Layout Engine

All layout primitives inherit from the base `Box` class. The layout engine is based on nested container components (`Row` and `Col`) that automatically calculate positions, margins, and merging (rowspan/colspan) for all child cells.

### Common Box Parameters

Every `Box` (including container and primitive nodes) supports the following layout parameters:

- `rowspan` (`int | None`): Number of rows the cell or container spans. If omitted, it's calculated automatically based on contents.
- `colspan` (`int | None`): Number of columns the cell or container spans. If omitted, calculated automatically.
- `offset` (`int`): Indentation (spacing) relative to the direction of layout. Defaults to `0`.
- `grow` (`bool`): If `True`, instructs this element to fill all remaining horizontal or vertical space within its parent. Defaults to `False`.

---

## Container Components

Containers hold and lay out other child components. They automatically handle spacing, alignment, and coordinate math.

### `Row`

A `Row` is a horizontal container that lays out its children left-to-right.

- **Children**: Lays out a list of child boxes horizontally.
- **Auto-span**: If a child does not have a `rowspan` specified, it automatically inherits the `rowspan` of the parent `Row`.

```python
from poi import Sheet, Row, Cell

sheet = Sheet(
    root=Row(
        rowspan=2,  # All children inherit a rowspan of 2
        children=[
            Cell("First Cell", colspan=2),
            Cell("Second Cell", offset=1),  # Offset leaves 1 column empty
        ]
    )
)
```

### `Col`

A `Col` is a vertical container that lays out its children top-to-bottom.

- **Children**: Lays out a list of child boxes vertically.
- **Auto-span**: If a child does not have a `colspan` specified, it automatically inherits the `colspan` of the parent `Col`.

```python
from poi import Sheet, Col, Cell

sheet = Sheet(
    root=Col(
        colspan=3,  # All children inherit a colspan of 3
        children=[
            Cell("Header Line"),
            Cell("Body Line", offset=1),  # Offset leaves 1 row empty
        ]
    )
)
```

---

## Primitive Components

Primitive nodes contain the actual content (data, images) written to the Excel file.

### `Cell`

A `Cell` represents a single rectangular merged region containing a specific value.

#### Parameters
- `value` (`CellValue`): The value of the cell. Supports `str`, `int`, `float`, `bool`, `datetime`, `date`, `time`, or `None`.
- `width` (`int | None`): If set, explicitly configures the width of the column(s) this cell resides in.
- `height` (`int | None`): If set, explicitly configures the height of the row this cell resides in.
- `comment` (`str | None`): Add an Excel cell comment/annotation.
- `comment_options` (`CommentOptions | None`): Configure comment layout and style (see [Styling Guide](styling.md)).
- `**kwargs`: Any formatting options from `CellStyle` (e.g. `bg_color`, `bold`, `align`, see [Styling Guide](styling.md)).

```python
Cell("Total:", bold=True, bg_color="#E0E0E0", align="right", comment="Calculated automatically")
```

### `Table`

A `Table` is a high-level component that automatically renders dynamic lists of objects or dictionaries with structured columns, headers, formatting, and dynamic conditional styles.

#### Parameters
- `data` (`Collection[T]`): List of objects (e.g., namedtuples, dicts, dataclasses).
- `columns` (`Collection[ColumnConfig]`): Column definitions (see below).
- `col_width` (`int | Literal["auto"] | None`): Standard column width for table columns. Defaults to `15`. Set to `"auto"` for automatic width adjustment.
- `row_height` (`RowHeightCallback[T] | int | None`): Height of rows in the table. Can be an integer or a callback function taking `(record, index)` for dynamic height.
- `cell_style` (`dict[str, RenderFunction[T]] | str | None`): CSS-like dynamic formatting mapping conditional functions to styles (see [Styling Guide](styling.md)).
- `date_format` (`str | None`): Number format pattern for `date` types (defaults to `yyyy-mm-dd`).
- `datetime_format` (`str | None`): Number format pattern for `datetime` types (defaults to `yyyy-mm-dd hh:mm:ss`).
- `time_format` (`str | None`): Number format pattern for `time` types (defaults to `hh:mm:ss`).
- `**kwargs`: Styles applied to the entire table (like `border`).

#### Column Configurations
Columns can be declared using two formats:
1. **Tuple format**: `(attribute_name, column_title)`. A quick shorthand for displaying attributes directly.
2. **Dictionary format**: Allows extensive customization:
    - `"title"` (`str`): The display header of the column (required).
    - `"attr"` (`str`): The attribute or nested dictionary key path to extract from data. Supports optional chaining (e.g., `order?.user?.name`).
    - `"render"` (`Callable`): A custom rendering function taking `(record, column)` to dynamically compute values.
    - `"width"` (`int | Literal["auto"] | None`): Specific column width.
    - `"type"` (`Literal["image", "text"]`): Set to `"image"` if rendering images.
    - `"options"` (`ImageOptions`): Options for images (scaling, offset).
    - `"format"` (`CellStyle`): Standard styling dictionary applied to cells in this column.
    - `"title_comment"` (`str`): Comment added to the table header.
    - `"title_comment_options"` (`CommentOptions`): Header comment options.

```python
columns = [
    # Tuple shorthand
    ("id", "ID"),
    # Advanced dictionary
    {
        "attr": "name",
        "title": "Product Name",
        "width": "auto",
        "title_comment": "Official display name"
    },
    # Custom rendering callback
    {
        "title": "Revenue",
        "render": lambda r, col: r.qty * r.price,
        "format": {"num_format": "$#,##0.00"}
    }
]
```

#### Column Auto-fit Features
Setting `width: "auto"` on specific columns, or `col_width="auto"` table-wide, enables dynamic auto-fitting:
- It scans all values in the column, including the title header, to find the longest string representation.
- CJK (Chinese, Japanese, Korean) characters are counted as double-width (2) for accurate sizing in Excel.
- Applies optimal column width with comfortable padding automatically, avoiding text clipping or `###` display issues.

---

### `Image`

Renders an image from a local path.

#### Parameters
- `filename` (`str`): Local system path to the image file.
- `options` (`ImageOptions | None`): Configure position, scaling, and tooltips:
    - `x_scale` (`float`): Horizontal scaling factor (e.g., `0.5`).
    - `y_scale` (`float`): Vertical scaling factor (e.g., `0.5`).
    - `x_offset` (`int`): Horizontal pixel offset.
    - `y_offset` (`int`): Vertical pixel offset.
    - `url` (`str`): Hyperlink attached to the image.
    - `tip` (`str`): Hover tooltip text.

```python
from poi import Sheet, Col, Image

sheet = Sheet(
    root=Col(
        children=[
            Image("assets/logo.png", options={"x_scale": 0.5, "y_scale": 0.5})
        ]
    )
)
```
