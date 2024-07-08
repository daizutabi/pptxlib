import pytest

from pptxlib.shapes import Shapes
from pptxlib.tables import Cell, CellRange, Columns, Rows, Table, Tables


def test_add(tables: Tables, shapes: Shapes):
    table = tables.add(2, 3, 100, 100, 240, 360)
    assert table.left == 100
    assert table.top == 100
    assert table.parent.width == 240
    assert table.parent.height == 360
    assert len(shapes) == 2
    assert len(tables) == 1
    table.delete()
    assert len(shapes) == 1
    assert len(tables) == 0


def test_call_getitem(tables: Tables):
    t1 = tables.add(2, 3, 100, 100, 200, 200)
    t2 = tables.add(4, 5, 200, 200, 300, 300)
    assert len(tables) == 2
    assert tables(1).left == 100
    assert tables(2).left == 200
    assert tables().left == 200
    assert tables[0].left == 100  # type: ignore
    assert tables[1].left == 200  # type: ignore
    t1.delete()
    t2.delete()


def test_left(table: Table):
    table.left = 10
    assert table.left == 10


def test_top(table: Table):
    table.top = 10
    assert table.top == 10


def test_rows(rows: Rows):
    assert len(rows) == 2


def test_columns(columns: Columns):
    assert len(columns) == 3


def test_shape(table: Table):
    assert table.shape == (2, 3)


@pytest.mark.parametrize("height", [None, 400])
def test_row_height(height, rows: Rows, table: Table):
    if height:
        table.height = height

    t = table.parent.height
    h = table.rows.height
    n = len(rows)
    for k, row in enumerate(rows):
        assert round(row.height) == round(t / n)
        assert row.height == h[k]


@pytest.mark.parametrize("width", [None, 400])
def test_column_widtht(width, columns: Columns, table: Table):
    if width:
        table.width = width

    t = table.parent.width
    w = table.columns.width
    n = len(columns)
    for k, column in enumerate(columns):
        assert round(column.width) == round(t / n)
        assert column.width == w[k]


def test_row_height_list(rows: Rows, table: Table):
    table.rows.height = [100, 200]
    assert rows(1).height == 100
    assert rows(2).height == 200
    assert table.parent.height == 300


def test_column_width_list(columns: Columns, table: Table):
    table.columns.width = [100, 200, 300]
    assert columns(1).width == 100
    assert columns(2).width == 200
    assert columns(3).width == 300
    assert table.parent.width == 600


def test_cell(table: Table):
    c1 = table.cell(2, 1)
    c2 = table.cell(4)
    assert c1.top == c2.top
    assert c1.left == c2.left
    assert c1.width == c2.width
    assert c1.height == c2.height

    c1 = table.cell(2, 3)
    c2 = table.cell(6)
    assert c1.top == c2.top
    assert c1.left == c2.left


def test_cell_text(table: Table):
    cell = table.cell(1, 1)
    cell.text = "a"
    assert cell.text == "a"
    cell.value = "a"
    assert cell.value == "a"


def test_minimize_height(table: Table):
    h1 = table.rows.height
    table.minimize_height()
    h2 = table.rows.height
    for x, y in zip(h1, h2, strict=True):
        assert x > y


def test_rows_cells(rows: Rows):
    cells = rows(1).cells
    assert cells.api.__class__.__name__ == "CellRange"
    assert cells.__class__.__name__ == "CellRange"
    assert len(cells) == 3
    assert len(rows) == 2


def test_colums_cells(columns: Columns):
    cells = columns(1).cells
    assert cells.api.__class__.__name__ == "CellRange"
    assert cells.__class__.__name__ == "CellRange"
    assert len(cells) == 2
    assert len(columns) == 3


def test_cell_from_cells(cell_range: CellRange):
    cell = cell_range(1)
    assert isinstance(cell, Cell)
    assert isinstance(cell.parent, Table)


def test_tables_repr(tables: Tables):
    assert repr(tables) == "<Tables>"


def test_table_repr(table: Table):
    assert repr(table) == "<Table [Table]>"


def test_rows_repr(rows: Rows):
    assert repr(rows) == "<Rows>"


def test_row_repr(rows: Rows):
    assert repr(rows(1)) == "<Row [Row]>"


def test_columns_repr(columns: Columns):
    assert repr(columns) == "<Columns>"


def test_column_repr(columns: Columns):
    assert repr(columns(1)) == "<Column [Column]>"


def test_cell_repr(table: Table):
    assert repr(table.cell(1)) == "<Cell [Cell]>"


def test_cells_repr(rows: Rows, columns: Columns):
    assert repr(rows(1).cells) == "<CellRange>"
    assert repr(columns(1).cells) == "<CellRange>"


def test_tables_parent(tables: Tables):
    assert tables.api.Parent.__class__.__name__ == "_Slide"
    assert tables.parent.__class__.__name__ == "Slide"


def test_table_parent(table: Table, tables: Tables):
    assert table.api.Parent.__class__.__name__ == "Shape"
    assert table.parent.__class__.__name__ == "Shape"
    assert tables(1).parent.__class__.__name__ == "Shape"


def test_rows_parent(rows: Rows):
    assert rows.api.Parent.__class__.__name__ == "Table"
    assert rows.parent.__class__.__name__ == "Table"


def test_row_parent(rows: Rows):
    assert rows(1).api.Parent.__class__.__name__ == "Table"
    assert rows(1).parent.__class__.__name__ == "Table"


def test_columns_parent(columns: Columns):
    assert columns.api.Parent.__class__.__name__ == "Table"
    assert columns.parent.__class__.__name__ == "Table"


def test_column_parent(columns: Columns):
    assert columns(1).api.Parent.__class__.__name__ == "Table"
    assert columns(1).parent.__class__.__name__ == "Table"


def test_cell_parent(cell: Cell):
    assert cell.api.Parent.__class__.__name__ == "Table"
    assert cell.parent.__class__.__name__ == "Table"


def test_cells_parent(rows: Rows, columns: Columns):
    assert rows(1).cells.api.Parent.__class__.__name__ == "Row"
    assert rows(1).cells.parent.__class__.__name__ == "Row"
    assert columns(1).cells.api.Parent.__class__.__name__ == "Column"
    assert columns(1).cells.parent.__class__.__name__ == "Column"


def test_cell_borders(cell: Cell):
    assert isinstance(cell, Cell)
    assert cell.borders.api.__class__.__name__ == "Borders"
    assert cell.borders.__class__.__name__ == "Borders"


def test_cell_range_borders(cell_range: CellRange):
    assert isinstance(cell_range, CellRange)
    assert cell_range.borders.api.__class__.__name__ == "Borders"
    assert cell_range.borders.__class__.__name__ == "Borders"


def test_cell_borders_parent(cell: Cell):
    assert cell.borders.api.Parent.__class__.__name__ == "Table"
    assert cell.borders.parent.__class__.__name__ == "Table"


def test_cell_range_borders_parent(cell_range: CellRange):
    assert isinstance(cell_range, CellRange)
    assert cell_range.borders.api.Parent.__class__.__name__ == "Table"
    assert cell_range.borders.parent.__class__.__name__ == "Table"


# def test_table_repr(table: Table):
#     assert repr(table) == "<Table [Table]>"


# def test_rows_repr(rows: Rows):
#     assert repr(rows) == "<Rows>"


# def test_row_repr(rows: Rows):
#     assert repr(rows(1)) == "<Row [Row]>"


# def test_columns_repr(columns: Columns):
#     assert repr(columns) == "<Columns>"


# def test_column_repr(columns: Columns):
#     assert repr(columns(1)) == "<Column [Column]>"


# a = 1

# from pptxlib import PowerPoint

# pp = PowerPoint()
# pr = pp.presentations.add()
# slide = pr.slides.add()
# table = slide.tables.add(2, 3, 10, 10)
# table.api.FirstRow = False
# table.api.HorizBanding = False

# cell = table.cell(1, 1)
# cell.set_border("left", 2)
