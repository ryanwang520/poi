import importlib.metadata

from .book import Book
from .nodes import (
    Alignment,
    BorderStyle,
    Box,
    Cell,
    CellStyle,
    CellValue,
    Col,
    Column,
    ColumnConfig,
    ColumnDict,
    ColumnTuple,
    CommentOptions,
    Image,
    ImageOptions,
    Row,
    Table,
    TableStyle,
    VerticalAlignment,
)
from .sheet import Sheet
from .writer import BytesIOWorkBook

# Main classes for public API
__all__ = [
    # Core classes
    "Sheet",
    "Book",
    "Cell",
    "Box",
    "Col",
    "Row",
    "Table",
    "Image",
    "BytesIOWorkBook",
    # Type definitions for enhanced typing
    "CellValue",
    "CellStyle",
    "CommentOptions",
    "ImageOptions",
    "TableStyle",
    # Column configuration types
    "Column",
    "ColumnDict",
    "ColumnTuple",
    "ColumnConfig",
    # Style literals
    "Alignment",
    "VerticalAlignment",
    "BorderStyle",
]

__version__ = importlib.metadata.version("poi")
