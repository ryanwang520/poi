# Styling & Comments

Poi provides rich styling, borders, and number formatting options that map directly to the formatting capabilities of Microsoft Excel. 

Additionally, Poi features a comprehensive comment and annotation system allowing you to attach customized pop-up notes to cells and headers.

---

## Styling Properties (`CellStyle`)

Styling is configured using standard keyword arguments on primitive components like `Cell` (or via the `format` dictionary in `Table` columns).

### Font Styles

- `font_name` (`str`): The font family name (e.g., `"Arial"`, `"Calibri"`, `"Segoe UI"`).
- `font_size` (`int`): Font size in points (e.g., `11`, `14`).
- `bold` (`bool`): Set to `True` for bold text.
- `italic` (`bool`): Set to `True` for italic text.
- `underline` (`bool`): Set to `True` to underline text.
- `font_color` (`str`): Hex code or name of the font color (e.g., `"#FF0000"` or `"red"`).

### Background & Fill Styles

- `bg_color` (`str`): Hex code or name of the background fill color (e.g., `"#D3D3D3"` or `"yellow"`).
- `pattern` (`int`): Fill pattern. `1` represents a solid fill (default when `bg_color` is set).

### Alignment & Layout

- `align` (`Alignment`): Horizontal cell alignment. Supports:
    - `"left"`
    - `"center"`
    - `"right"`
    - `"justify"`
- `valign` (`VerticalAlignment`): Vertical cell alignment. Supports:
    - `"top"`
    - `"vcenter"` (vertical center)
    - `"bottom"`
    - `"vjustify"`
- `text_wrap` (`bool`): Set to `True` to enable wrapping of long text in the cell.
- `shrink_to_fit` (`bool`): Set to `True` to shrink font size automatically so text fits the column width.
- `indent` (`int`): Horizontal cell indentation level.
- `rotation` (`int`): Text rotation angle in degrees (e.g., `90` or `-45`).

---

## Borders

Excel supports a variety of border styles, specified as integers from `0` to `6` (mapping to xlsxwriter's border indices):

- `0`: No border
- `1`: Thin border
- `2`: Medium border
- `3`: Dashed border
- `4`: Dotted border
- `5`: Thick border
- `6`: Double border

### Border Configuration Keys

- `border` (`int`): Apply the specified border style to all four sides of the cell.
- `border_color` (`str`): Color of all borders (e.g., `"#808080"`).
- `top` (`int`), `bottom` (`int`), `left` (`int`), `right` (`int`): Configure specific borders individually.
- `top_color` (`str`), `bottom_color` (`str`), `left_color` (`str`), `right_color` (`str`): Set individual border colors.

```python
# Cell with a thick bottom border and thin side borders
Cell("Header", bottom=5, bottom_color="#000000", left=1, right=1)
```

---

## Number Formatting (`num_format`)

Excel uses format strings to render numbers, currencies, percentages, and dates properly:

- `num_format` (`str`): Custom format string (e.g., `"$#,##0.00"`, `"0.0%"`, `"yyyy-mm-dd"`).

```python
# Format as local currency with commas and 2 decimal points
Cell(1250000.50, num_format="$#,##0.00")  # Displays as: $1,250,000.50

# Format as percentage
Cell(0.854, num_format="0.0%")  # Displays as: 85.4%
```

---

## Dynamic Table Styles (`cell_style`)

In a `Table`, you can apply conditional styles dynamically using a mapping of CSS-like strings to lambda functions. The lambda takes `(record, column)` and applies the style if the condition evaluates to `True`.

```python
from poi import Sheet, Table

sheet = Sheet(
    root=Table(
        data=data,
        columns=columns,
        cell_style={
            # Apply bold red font if price is over 100
            "font_color: red; bold: true": lambda r, col: col.attr == "price" and r.price > 100,
            # Apply light gray background to name column
            "bg_color: #F5F5F5": lambda r, col: col.attr == "name"
        }
    )
)
```

---

## Cell Comments & Annotations

Both individual cells and table headers support attaching custom Excel comments (tooltips).

### `CommentOptions`

Configure the appearance and behavior of cell comments with these options:

- `author` (`str`): The name of the comment author.
- `visible` (`bool`): If `True`, the comment is displayed continuously. If `False` (default), it appears only on hover.
- `color` (`str`): Hex background color of the comment box (e.g., `"#FFFFE0"`).
- `x_scale` (`float`): Horizontal scale factor of the comment box.
- `y_scale` (`float`): Vertical scale factor of the comment box.
- `width` (`int`): Custom width of the comment box in pixels.
- `height` (`int`): Custom height of the comment box in pixels.

### Adding Comments to a `Cell`

Pass `comment` and `comment_options` arguments to `Cell`:

```python
from poi import Cell

Cell(
    "Key Metric",
    comment="This metric is calculated based on monthly active users.",
    comment_options={
        "author": "Analytics Team",
        "color": "#E6F2FF",
        "x_scale": 1.5,
        "y_scale": 1.2
    }
)
```

### Adding Comments to Table Headers

Provide `"title_comment"` and optional `"title_comment_options"` in the dictionary column definition:

```python
columns = [
    ("id", "ID"),
    {
        "attr": "price",
        "title": "Unit Price",
        "title_comment": "Price in USD. Includes regional sales tax.",
        "title_comment_options": {
            "author": "Finance Team",
            "visible": True  # Always visible in the spreadsheet
        }
    }
]
```
