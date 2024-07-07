import pytest

from pptxlib.shapes import Shapes
from pptxlib.tables import Columns, Rows, Table, Tables


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


def test_repr_tables(tables: Tables):
    assert repr(tables) == "<Tables>"


def test_repr_table(table: Table):
    assert repr(table) == "<Table [Table]>"


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
