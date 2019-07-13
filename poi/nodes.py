from typing import Optional, Tuple

from poi.helpers import Direction


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
            self.box.add_child_span(child)
            child_node, current_row, current_col = self.box.process_child(
                child, current_row, current_col
            )
            self.children.append(child_node)


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
            return Direction.HORIZONTAL
        if isinstance(self.parent, Col):
            return Direction.VERTICAL
        raise ValueError(f"invalid parent {self.parent}")

    @property
    def is_horizontal(self):
        return self.direction == Direction.HORIZONTAL

    @property
    def is_vertical(self):
        return self.direction == Direction.VERTICAL

    def __init__(
        self, children=None, rowspan=None, colspan=None, offset=0, grow=False, **kwargs
    ):
        self.parent = None
        self.rowspan = rowspan
        self.colspan = colspan
        self.offset = offset
        self.grow = grow
        self.children = children or []
        for child in self.children:
            child.parent = self
        self.styles = kwargs
        self.instance = None

    def add_child_span(self, child):
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
        assert all(child.parent for child in self.children)


class Row(Box):
    def process_child(
        self, child, current_row, current_col
    ) -> Tuple["BoxInstance", int, int]:
        child_node = BoxInstance(
            child, current_row, current_col + child.offset, self.instance
        )
        current_col += child.cols
        return child_node, current_row, current_col

    def add_child_span(self, child):
        if not child.rowspan:
            child.rowspan = self.rowspan
        if child.grow:
            child.colspan = self.colspan - self.offset

    @property
    def cols(self):
        """
        cols and row can only be accessed if all chilren are bound to instance
        """
        if self.colspan:
            offset = self.offset if self.is_horizontal else 0
            return self.colspan + offset
        self.assert_children_bound()
        return sum(child.cols for child in self.children)

    @property
    def rows(self):
        if self.rowspan:
            offset = self.offset + self.offset if self.is_vertical else 0
            return self.rowspan + offset
        self.assert_children_bound()
        return max(child.rows for child in self.children)


class Col(Box):
    def add_child_span(self, child):
        if not child.colspan:
            child.colspan = self.colspan
        if child.grow:
            child.rowspan = self.rowspan - self.offset

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


class Cell(Box):
    def __init__(self, value, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.value = value

    @property
    def cols(self):
        offset = self.offset if self.is_horizontal else 0
        return (self.colspan or 1) + offset

    @property
    def rows(self):
        offset = self.offset if self.is_vertical else 0
        return (self.rowspan or 1) + offset


class Image(Cell):
    pass


class Table(Box):
    def __init__(
        self,
        data,
        headers,
        cell_width=None,
        cell_style=None,
        date_format=None,
        *args,
        **kwargs,
    ):
        super().__init__(*args, **kwargs)
        self.data = data
        self.headers = headers
        self.cell_width = cell_width
        self.cell_style = cell_style or {}
        self.date_format = date_format
        self.attr_map = {name: val for name, val in headers}

    @property
    def rows(self):
        return len(self.data) + 1

    @property
    def cols(self,):
        return len(self.headers)
