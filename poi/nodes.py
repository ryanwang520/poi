from __future__ import annotations

import logging
from collections import abc
from datetime import date, datetime, time
from typing import (
    Any,
    Callable,
    Collection,
    Generic,
    Iterable,
    NamedTuple,
    Protocol,
    TypedDict,
    TypeVar,
    Union,
)

from typing_extensions import Literal, NotRequired, Unpack

logger = logging.getLogger("poi")
logger.addHandler(logging.NullHandler())

# =================== TYPE DEFINITIONS ===================

# Basic type aliases
CellValue = Union[str, int, float, bool, datetime, date, time, None]
Direction = Literal["HORIZONTAL", "VERTICAL"]

# Style and formatting types
Alignment = Literal["left", "center", "right", "justify"]
VerticalAlignment = Literal["top", "vcenter", "bottom", "vjustify"]
BorderStyle = Literal[0, 1, 2, 3, 4, 5, 6]  # xlsxwriter border styles


class CommentOptions(TypedDict):
    """Options for Excel cell comments."""

    author: NotRequired[str]
    visible: NotRequired[bool]
    x_scale: NotRequired[float]
    y_scale: NotRequired[float]
    color: NotRequired[str]
    width: NotRequired[int]
    height: NotRequired[int]


class CellStyle(TypedDict):
    """Comprehensive cell style options."""

    # Font styling
    font_name: NotRequired[str]
    font_size: NotRequired[int]
    bold: NotRequired[bool]
    italic: NotRequired[bool]
    underline: NotRequired[bool]
    font_color: NotRequired[str]

    # Cell styling
    bg_color: NotRequired[str]
    pattern: NotRequired[int]
    align: NotRequired[Alignment]
    valign: NotRequired[VerticalAlignment]
    text_wrap: NotRequired[bool]
    shrink_to_fit: NotRequired[bool]
    indent: NotRequired[int]
    rotation: NotRequired[int]

    # Border styling
    border: NotRequired[BorderStyle]
    border_color: NotRequired[str]
    top: NotRequired[BorderStyle]
    bottom: NotRequired[BorderStyle]
    left: NotRequired[BorderStyle]
    right: NotRequired[BorderStyle]
    top_color: NotRequired[str]
    bottom_color: NotRequired[str]
    left_color: NotRequired[str]
    right_color: NotRequired[str]

    # Number formatting
    num_format: NotRequired[str]


class ImageOptions(TypedDict):
    """Options for image insertion."""

    x_scale: NotRequired[float]
    y_scale: NotRequired[float]
    x_offset: NotRequired[int]
    y_offset: NotRequired[int]
    object_position: NotRequired[int]
    url: NotRequired[str]
    tip: NotRequired[str]
    description: NotRequired[str]


class TableStyle(TypedDict):
    """Style options specific to tables."""

    border: NotRequired[BorderStyle]
    align: NotRequired[Alignment]
    valign: NotRequired[VerticalAlignment]
    bg_color: NotRequired[str]
    font_size: NotRequired[int]
    bold: NotRequired[bool]
    text_wrap: NotRequired[bool]


# Protocol definitions for callbacks
T_contra = TypeVar("T_contra", contravariant=True)


# Render function types for column configuration
RenderFunction = Union[
    Callable[[], CellValue],
    Callable[[T_contra], CellValue],  # Simple: lambda record: record.field
    Callable[
        [T_contra, "Column"], CellValue
    ],  # Advanced: lambda record, column: f"{record.field} ({column.title})"
]


class StyleCondition(Protocol[T_contra]):
    """Protocol for conditional style functions."""

    def __call__(self, record: T_contra, column: Column) -> bool: ...


class RowHeightCallback(Protocol[T_contra]):
    """Protocol for dynamic row height functions."""

    def __call__(self, record: T_contra, index: int, *args: Any) -> int: ...


# Column configuration types
class ColumnDict(TypedDict):
    """Dictionary format for column configuration."""

    title: str
    attr: NotRequired[str]
    render: NotRequired[RenderFunction[Any]]
    width: NotRequired[int]
    type: NotRequired[Literal["image", "text"]]
    options: NotRequired[ImageOptions]
    format: NotRequired[CellStyle]
    title_comment: NotRequired[str]
    title_comment_options: NotRequired[CommentOptions]


ColumnTuple = tuple[str, str]  # (attr, title)
ColumnConfig = Union[ColumnDict, ColumnTuple]

# Generic type variable with constraint
T = TypeVar("T")
TableRow = TypeVar("TableRow", bound=Any)


class BoxInstance:
    def __init__(
        self, box: Box, row: int, col: int, parent: BoxInstance | None
    ) -> None:
        self.box = box
        self.row = row
        self.col = col
        self.parent = parent
        box.instance = self
        self._setup_children()

    def _setup_children(self) -> None:
        children = self.box.children
        if not children:
            return
        children = [
            Cell(child) if isinstance(child, str) else child for child in children
        ]

        self.children = []

        current_row = self.row
        current_col = self.col
        for child in children:
            child.styles = {**child.styles, **self.box.styles}
            self.box.add_child_span(
                child, neighbours=[c for c in children if c is not child]
            )
            child_node, current_row, current_col = self.box.layout_child_element(
                child, current_row, current_col
            )
            self.children.append(child_node)


_NotDetermined = object()


Visitor = Callable[["Box"], None]


class Box:
    instance: BoxInstance
    parent: Box | None
    styles: CellStyle

    def accept(self, visitor: Visitor) -> None:
        visitor(self)

    @property
    def row(self) -> int:
        assert self.instance
        return self.instance.row

    @property
    def col(self) -> int:
        assert self.instance
        return self.instance.col

    @property
    def direction(self) -> Direction:
        if not self.parent:
            raise ValueError("root node cannot have direction")
        if isinstance(self.parent, Row):
            return "HORIZONTAL"
        if isinstance(self.parent, Col):
            return "VERTICAL"
        raise ValueError(f"invalid parent {self.parent}")

    @property
    def is_horizontal(self) -> bool:
        return self.direction == "HORIZONTAL"

    @property
    def is_vertical(self) -> bool:
        return self.direction == "VERTICAL"

    children: list[Box]

    def __init__(
        self,
        children: Iterable[Box | None] | Box | None = None,
        rowspan: int | None = None,
        colspan: int | None = None,
        offset: int = 0,
        grow: bool = False,
        **kwargs: Unpack[CellStyle],
    ) -> None:
        self.parent = None
        self.rowspan = rowspan
        self.colspan = colspan
        self.offset = offset
        self.grow = grow

        def flatten(items: Any) -> Iterable[Any]:
            """Yield items from any nested iterable."""
            for x in items:
                if isinstance(x, abc.Iterable) and not isinstance(x, (str, bytes)):
                    yield from flatten(x)
                else:
                    yield x

        if isinstance(children, Box):
            children = [children]
        else:
            children = [child for child in flatten(children or []) if child is not None]
        self.children = children  # type: ignore
        for child in self.children:
            child.parent = self
        self.styles = kwargs
        self.instance = None  # type: ignore

    def add_child_span(self, child: Box, neighbours: list[Box]) -> None:
        pass

    @property
    def cell_format(self) -> CellStyle:
        return self.styles or {}

    def layout_child_element(
        self, child: Box, current_row: int, current_col: int
    ) -> Any:
        raise ValueError(f"not support type for {self}")

    def __repr__(self) -> str:
        if not self.instance:
            return f"""
        Unbound box {self.__class__.__name__}
        rowspan {self.rowspan}
        colspan {self.colspan}
            """
        return f"""
        {self.__class__.__name__}
        row {self.row}
        col {self.col}
        rowspan {self.rowspan}
        colspan {self.colspan}
        """

    @property
    def cols(self) -> int:
        raise NotImplementedError

    @property
    def rows(self) -> int:
        raise NotImplementedError

    def assert_children_bound(self) -> None:
        assert all(child.instance for child in self.children)

    def calculate_column_span(self, raises: bool = True) -> Any:
        if self.colspan:
            return self.colspan + self.offset
        if self.grow:
            if raises:
                raise ValueError(f"{self} cannot grow as the col have to be determined")
            return _NotDetermined
        max_col = 1
        for child in self.children:
            if child.colspan:
                if max_col < child.colspan:
                    max_col = child.colspan
            else:
                col = child.calculate_column_span(raises=raises)
                if col is not _NotDetermined and col > max_col:
                    max_col = col
        return max_col + self.offset

    def calculate_row_span(self, raises: bool = True) -> Any:
        if self.rowspan:
            return self.rowspan + self.offset
        if self.grow:
            if raises:
                raise ValueError(f"{self} cannot grow as the row have to be determined")
            return _NotDetermined
        max_row = 1
        for child in self.children:
            if child.rowspan:
                if max_row < child.rowspan:
                    max_row = child.rowspan
            else:
                row = child.calculate_row_span(raises=raises)
                if row is not _NotDetermined and row > max_row:
                    max_row = row
        return max_row + self.offset


class Row(Box):
    colspan: int

    def layout_child_element(
        self, child: Box, current_row: int, current_col: int
    ) -> tuple[BoxInstance, int, int]:
        child_node = BoxInstance(
            child, current_row, current_col + child.offset, self.instance
        )
        current_col += child.cols
        return child_node, current_row, current_col

    def add_child_span(self, child: Box, neighbours: list[Box]) -> None:
        if not child.rowspan:
            child.rowspan = self.rowspan
        if child.grow:
            assert all(not c.grow for c in neighbours), (
                "only one col in a row can have grow attr"
            )
            if not self.colspan:
                if self.instance.parent:
                    parent = self.instance.parent.box
                else:
                    raise ValueError(f"{child} width is not determinable")

                neighbor_with_cols = [child for child in parent.children if child]
                if neighbor_with_cols:
                    neighbour_cols = [
                        n.calculate_column_span(raises=False)
                        for n in neighbor_with_cols
                    ]
                    valid_cols = [
                        col for col in neighbour_cols if col is not _NotDetermined
                    ]
                    if valid_cols:
                        self.colspan = max(valid_cols)
                    else:
                        raise ValueError(f"{child} width is not determinable")
                else:
                    raise ValueError(f"{child} width is not determinable")
            child.colspan = (
                self.colspan
                - child.offset
                - sum(c.calculate_column_span() for c in neighbours)
            )

    @property
    def cols(self) -> int:
        """Cols and row can only be accessed if all children are bound to instance."""
        offset = self.offset if self.is_horizontal else 0
        if self.colspan:
            return self.colspan + offset
        self.assert_children_bound()
        return sum(child.cols for child in self.children) + offset

    @property
    def rows(self) -> int:
        offset = self.offset if self.is_vertical else 0
        if self.rowspan:
            return self.rowspan + offset
        self.assert_children_bound()
        return max(child.rows for child in self.children) + offset


class Col(Box):
    rowspan: int

    def add_child_span(self, child: Box, neighbours: list[Box]) -> None:
        if not child.colspan:
            child.colspan = self.colspan
        if child.grow:
            assert all(not c.grow for c in neighbours), (
                "only one row in a col can have grow attr"
            )
            if not self.rowspan:
                if self.instance.parent:
                    parent = self.instance.parent.box
                else:
                    raise ValueError(f"{child} height is not determinable")
                neighbor_with_rows = [child for child in parent.children if child]
                if neighbor_with_rows:
                    neighbour_rows = [
                        n.calculate_row_span(raises=False) for n in neighbor_with_rows
                    ]
                    valid_rows = [
                        row for row in neighbour_rows if row is not _NotDetermined
                    ]
                    if valid_rows:
                        self.rowspan = max(valid_rows)
                    else:
                        raise ValueError(f"{child} height is not determinable")
                else:
                    raise ValueError(f"{child} height is not determinable")
            child.rowspan = (
                self.rowspan
                - child.offset
                - sum(c.calculate_row_span() for c in neighbours)
            )

    def layout_child_element(
        self, child: Box, current_row: int, current_col: int
    ) -> tuple[BoxInstance, int, int]:
        child_node = BoxInstance(
            child, current_row + child.offset, current_col, self.instance
        )
        current_row += child.rows
        return child_node, current_row, current_col

    @property
    def cols(self) -> int:
        if self.colspan:
            offset = self.offset if self.is_horizontal else 0
            return self.colspan + offset
        self.assert_children_bound()
        return max(child.cols for child in self.children)

    @property
    def rows(self) -> int:
        if self.rowspan:
            offset = self.offset if self.is_vertical else 0
            return self.rowspan + offset
        self.assert_children_bound()
        return sum(child.rows for child in self.children)


class PrimitiveBox(Box):
    @property
    def cols(self) -> int:
        offset = self.offset if self.is_horizontal else 0
        return (self.colspan or 1) + offset

    @property
    def rows(self) -> int:
        offset = self.offset if self.is_vertical else 0
        return (self.rowspan or 1) + offset


class Cell(PrimitiveBox):
    def __init__(
        self,
        value: CellValue,
        *,  # Force keyword-only arguments
        # Box layout parameters
        rowspan: int | None = None,
        colspan: int | None = None,
        offset: int = 0,
        grow: bool = False,
        # Cell-specific parameters
        width: int | None = None,
        height: int | None = None,
        comment: str | None = None,
        comment_options: CommentOptions | None = None,
        # Formatting parameters from CellStyle
        **kwargs: Unpack[CellStyle],
    ) -> None:
        super().__init__(
            children=None,  # Cell is primitive, no children
            rowspan=rowspan,
            colspan=colspan,
            offset=offset,
            grow=grow,
            **kwargs,
        )
        self.value = value
        self.height = height
        self.width = width
        self.comment = comment
        self.comment_options = comment_options or {}


class Image(PrimitiveBox):
    def __init__(
        self,
        filename: str,
        *,  # Force keyword-only arguments
        # Box layout parameters
        rowspan: int | None = None,
        colspan: int | None = None,
        offset: int = 0,
        grow: bool = False,
        # Image-specific parameters
        options: ImageOptions | None = None,
    ) -> None:
        super().__init__(
            children=None,  # Image is primitive, no children
            rowspan=rowspan,
            colspan=colspan,
            offset=offset,
            grow=grow,
        )
        self.filename = filename
        self.options = options


class Column(NamedTuple):
    """Configuration for table columns."""

    title: str
    attr: str | None = None
    render: RenderFunction[Any] | None = None
    width: int | None = None
    type: Literal["image", "text"] = "text"
    options: ImageOptions | None = None
    format: CellStyle | None = None
    title_comment: str | None = None
    title_comment_options: CommentOptions | None = None


class Table(Box, Generic[T]):
    """High-level component for tabular data with headers and formatting."""

    columns: list[Column]

    def __init__(
        self,
        data: Collection[T],
        columns: Collection[ColumnConfig],
        col_width: int | None = None,
        row_height: RowHeightCallback[T] | int | None = None,
        cell_style: dict[str, RenderFunction[T]] | str | None = None,
        datetime_format: str | None = None,
        date_format: str | None = None,
        time_format: str | None = None,
        # Table-wide style parameters (includes border)
        **kwargs: Unpack[CellStyle],
    ) -> None:
        # Set default border if not provided
        kwargs.setdefault("border", 1)
        super().__init__(
            children=None,  # Table manages its own children
            **kwargs,
        )
        self.data = data
        self.col_width = col_width or 15
        self.row_height = row_height

        self.cell_style = cell_style or {}
        self.date_format = date_format
        self.datetime_format = datetime_format
        self.time_format = time_format
        self.columns = []
        for col in columns:
            if isinstance(col, tuple):
                # Only support 2-tuple: (attr, title)
                if len(col) != 2:
                    raise ValueError(
                        f"Tuple must have exactly 2 elements (attr, title), "
                        f"got {len(col)}"
                    )
                item = Column(attr=col[0], title=col[1])
            elif isinstance(col, dict):
                item = Column(
                    attr=col.get("attr"),
                    title=col["title"],
                    type=col.get("type"),  # type: ignore
                    options=col.get("options"),
                    render=col.get("render"),
                    width=col.get("width"),
                    format=col.get("format"),
                    title_comment=col.get("title_comment"),
                    title_comment_options=col.get("title_comment_options"),
                )
            else:
                raise ValueError(f"Column must be tuple or dict, got {type(col)}")
            self.columns.append(item)

        self.rowspan = len(self.data) + 1
        self.colspan = len(self.columns)

    @property
    def rows(self) -> int:
        offset = self.offset if self.is_vertical else 0
        return self.rowspan + offset  # type: ignore

    @property
    def cols(
        self,
    ) -> int:
        offset = self.offset if self.is_horizontal else 0
        assert self.colspan
        return self.colspan + offset
