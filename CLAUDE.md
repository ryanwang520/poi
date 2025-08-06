# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Common Development Commands

### Testing
- `uv run pytest` - Run all tests
- `uv run pytest --update-snapshot` - Run tests and update snapshot files (for Excel file comparisons)
- `uv run pytest tests/test_specific.py` - Run specific test file

### Code Quality
- `uv run ruff check` - Run linting (configured in pyproject.toml)
- `uv run ruff format` - Format code
- `uv run mypy poi` - Run type checking (configured in mypy.ini, excludes tests)

### Documentation
- `uv run mkdocs serve` - Serve documentation locally for development
- `uv run mkdocs build` - Build documentation
- `uv run mkdocs gh-deploy -r upstream` - Deploy documentation to GitHub Pages

### Building and Publishing
- `uv build` - Build the package using uv_build backend
- `uv publish` - Publish to PyPI (after building)

## Architecture Overview

Poi is a declarative Excel generation library built around a layout engine with visitor pattern for writing Excel files.

### Core Components

**Layout System (`poi/nodes.py`)**
- `Box` - Base class for all layout elements with positioning, spanning, and styling
- `Row` - Horizontal container that layouts children left-to-right 
- `Col` - Vertical container that layouts children top-to-bottom
- `Cell` - Primitive element containing a single value
- `Table` - High-level component for tabular data with headers, formatting, and conditional styling
- `Image` - Primitive for embedding images in Excel files

**Sheet Management (`poi/sheet.py`, `poi/book.py`)**
- `Sheet` - Main interface for creating single worksheet Excel files from a root Box
- `Book` - Container for multiple sheets to create multi-worksheet Excel files

**Visitor Pattern (`poi/visitors/`)**
- `writer_visitor` - Traverses the Box tree and writes to Excel using xlsxwriter
- `print_visitor` - Debug visitor for printing the layout structure
- Writers handle Excel-specific formatting, merging, images, and conditional styling

### Key Design Patterns

1. **Declarative Layout** - Define Excel structure using nested Box components rather than imperative cell-by-cell writing
2. **Visitor Pattern** - Separate layout logic from Excel writing logic for extensibility
3. **Auto Layout** - Boxes automatically calculate their size and position based on children and constraints
4. **Span Management** - Automatic rowspan/colspan calculation with support for growing elements
5. **Conditional Styling** - CSS-like style expressions with lambda conditions for dynamic formatting

### Test Structure

Tests use snapshot testing comparing generated Excel file sizes (within tolerance) against reference files in `tests/__snapshots__/`. The test framework supports updating snapshots when Excel structure changes.

### Dependencies

- `xlsxwriter` - Core Excel file generation
- Development: `pytest`, `ruff`, `mypy`, `mkdocs` with Material theme
- Build system uses `uv_build` backend with Python 3.8+ support