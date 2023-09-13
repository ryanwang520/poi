from functools import singledispatch
from typing import Any

from ..nodes import Cell, Col, Image, Row, Table


@singledispatch
def print_visitor(_: Any) -> None:
    pass


@print_visitor.register
def _(self: Col) -> None:
    print(f"write Col at {self.row}:{self.col}")
    for child in self.children:
        print_visitor(child)


@print_visitor.register  # type: ignore
def _(self: Row):
    print(f"write Row at {self.row}:{self.col}")
    for child in self.children:
        print_visitor(child)


@print_visitor.register  # type: ignore
def _(self: Table):  # type: ignore
    print(f"write Table at {self.row}:{self.col}")


@print_visitor.register  # type: ignore
def _(self: Cell):
    print(
        f"write Cell {self.value} at {self.row}:{self.col} "
        f"rowspan {self.rowspan} colspan {self.colspan}"
    )


@print_visitor.register  # type: ignore
def _(self: Image):
    print(f"insert image {self.filename} at {self.row}:{self.col}")
