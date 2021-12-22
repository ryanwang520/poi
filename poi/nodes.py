from collections import abc
from typing import (
    Optional,
    Tuple,
    NamedTuple,
    Callable,
    Dict,
    Union,
    TypeVar,
    Generic,
    Iterable,
    Collection,
    Any,
    List,
)

try:
    from typing import Literal
except ImportError:
    from typing_extensions import Literal  # type: ignore

import logging

logger = logging.getLogger("poi")
logger.addHandler(logging.NullHandler())

Direction = Literal["HORIZONTAL", "VERTICAL"]


class BoxInstance:
    def __init__(self, box: "Box", row, col, parent):
        self.box = box
        self.row = row
        self.col = col
        self.parent = parent
        box.instance = self
        self._setup_children()

    def _setup_children(self):
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
            child_node, current_row, current_col = self.box.process_child(
                child, current_row, current_col
            )
            self.children.append(child_node)


_NotDetermined = object()


class Box:
    instance: Optional[BoxInstance]
    parent: Optional["Box"]

    def accept(self, visitor):
        visitor(self)

    @property
    def row(self):
        assert self.instance
        return self.instance.row

    @property
    def col(self):
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

    def __init__(
        self,
        children: Union[Iterable[Optional["Box"]], "Box"] = None,
        rowspan: int = None,
        colspan: int = None,
        offset: int = 0,
        grow: bool = False,
        **kwargs,
    ):
        self.parent = None
        self.rowspan = rowspan
        self.colspan = colspan
        self.offset = offset
        self.grow = grow

        def flatten(items):
            """Yield items from any nested iterable"""
            for x in items:
                if isinstance(x, abc.Iterable) and not isinstance(x, (str, bytes)):
                    for sub_x in flatten(x):
                        yield sub_x
                else:
                    yield x

        if isinstance(children, Box):
            children = [children]
        else:
            children = [child for child in flatten(children or []) if child is not None]
        self.children = children
        for child in self.children:
            child.parent = self  # type: ignore
        self.styles = kwargs
        self.instance = None

    def add_child_span(self, child, neighbours):
        pass

    @property
    def cell_format(self):
        styles = self.styles or {}
        return styles

    def process_child(self, child, current_row, current_col):
        raise ValueError(f"not support type for {self}")

    def __repr__(self):
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
    def cols(self):
        raise NotImplementedError

    @property
    def rows(self):
        raise NotImplementedError

    def assert_children_bound(self):
        assert all(child.instance for child in self.children)

    def figure_out_cols(self, raises=True):
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
                col = child.figure_out_cols(raises=raises)
                if col is not _NotDetermined and col > max_col:
                    max_col = col
        return max_col + self.offset

    def figure_out_rows(self, raises=True):
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
                row = child.figure_out_rows(raises=raises)
                if row is not _NotDetermined and row > max_row:
                    max_row = row
        return max_row + self.offset


class Row(Box):
    def process_child(
        self, child, current_row, current_col
    ) -> Tuple["BoxInstance", int, int]:
        child_node = BoxInstance(
            child, current_row, current_col + child.offset, self.instance
        )
        current_col += child.cols
        return child_node, current_row, current_col

    def add_child_span(self, child, neighbours):
        if not child.rowspan:
            child.rowspan = self.rowspan
        if child.grow:
            assert all(
                not c.grow for c in neighbours
            ), "only one col in a row can have grow attr"
            if not self.colspan:
                if self.instance.parent:
                    parent = self.instance.parent.box
                else:
                    raise ValueError(f"{child} width is not determinable")

                neighbor_with_cols = [child for child in parent.children if child]
                if neighbor_with_cols:
                    neighbour_cols = [
                        n.figure_out_cols(raises=False) for n in neighbor_with_cols
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
                - sum(c.figure_out_cols() for c in neighbours)
            )

    @property
    def cols(self):
        """
        cols and row can only be accessed if all chilren are bound to instance
        """
        offset = self.offset if self.is_horizontal else 0
        if self.colspan:
            return self.colspan + offset
        self.assert_children_bound()
        return sum(child.cols for child in self.children) + offset

    @property
    def rows(self):
        offset = self.offset if self.is_vertical else 0
        if self.rowspan:
            return self.rowspan + offset
        self.assert_children_bound()
        return max(child.rows for child in self.children) + offset


class Col(Box):
    def add_child_span(self, child, neighbours):
        if not child.colspan:
            child.colspan = self.colspan
        if child.grow:
            assert all(
                not c.grow for c in neighbours
            ), "only one row in a col can have grow attr"
            if not self.rowspan:

                if self.instance.parent:
                    parent = self.instance.parent.box
                else:
                    raise ValueError(f"{child} height is not determinable")
                neighbor_with_rows = [child for child in parent.children if child]
                if neighbor_with_rows:
                    neighbour_rows = [
                        n.figure_out_rows(raises=False) for n in neighbor_with_rows
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
                - sum(c.figure_out_rows() for c in neighbours)
            )

    def process_child(self, child, current_row, current_col):
        child_node = BoxInstance(
            child, current_row + child.offset, current_col, self.instance
        )
        current_row += child.rows
        return child_node, current_row, current_col

    @property
    def cols(self):
        if self.colspan:
            offset = self.offset if self.is_horizontal else 0
            return self.colspan + offset
        self.assert_children_bound()
        return max(child.cols for child in self.children)

    @property
    def rows(self):
        if self.rowspan:
            offset = self.offset if self.is_vertical else 0
            return self.rowspan + offset
        self.assert_children_bound()
        return sum(child.rows for child in self.children)


class PrimitiveBox(Box):
    @property
    def cols(self):
        offset = self.offset if self.is_horizontal else 0
        return (self.colspan or 1) + offset

    @property
    def rows(self):
        offset = self.offset if self.is_vertical else 0
        return (self.rowspan or 1) + offset


class Cell(PrimitiveBox):
    def __init__(self, value: str, *args, **kwargs):
        height = kwargs.pop("height", None)
        super().__init__(*args, **kwargs)
        self.value = value
        self.height = height


class Image(PrimitiveBox):
    def __init__(self, filename: str, *args, **kwargs):
        options = kwargs.pop("options", None)
        super().__init__(*args, **kwargs)
        self.filename = filename
        self.options = options


class Column(NamedTuple):
    title: str
    attr: Optional[str] = None
    render: Optional[Callable] = None
    width: Optional[int] = None
    type: Literal["image", "text"] = "text"
    options: Optional[dict] = None
    format: Optional[dict] = None


T = TypeVar("T")


class Table(Box, Generic[T]):
    columns: List[Column]

    def __init__(
        self,
        data: Collection[T],
        columns: Collection[Any],
        cell_width: int = 15,
        col_width: int = 15,
        cell_height: int = 20,
        row_height: Union[Callable, int] = None,
        border: int = 1,
        cell_style: Union[
            Dict[str, Union[Callable[[T, Column], bool], Callable[[T], bool]]], str
        ] = None,
        datetime_format: str = None,
        date_format: str = None,
        time_format: str = None,
        *args,
        **kwargs,
    ):
        super().__init__(*args, border=border, **kwargs)
        self.data = data
        if cell_width:
            logger.warning("cell_width is deprecated, use col_width instead")
        self.col_width = col_width or cell_width
        if cell_height:
            logger.warning("cell_height is deprecated, use row_height instead")
        self.row_height = row_height or cell_height

        self.cell_style = cell_style or {}
        self.date_format = date_format
        self.datetime_format = datetime_format
        self.time_format = time_format
        self.columns = []
        for col in columns:
            assert isinstance(col, (tuple, dict))
            if isinstance(col, tuple):
                item = Column(attr=col[0], title=col[1])
            else:
                item = Column(
                    attr=col.get("attr"),
                    title=col["title"],
                    type=col.get("type"),  # type: ignore
                    options=col.get("options"),
                    render=col.get("render"),
                    width=col.get("width"),
                    format=col.get('format'),
                )
            self.columns.append(item)

        self.rowspan = len(self.data) + 1
        self.colspan = len(self.columns)

    @property
    def rows(self):
        offset = self.offset if self.is_vertical else 0
        return self.rowspan + offset

    @property
    def cols(self,) -> int:
        offset = self.offset if self.is_horizontal else 0
        assert self.colspan
        return self.colspan + offset
