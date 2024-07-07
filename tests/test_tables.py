from win32com.client import DispatchBaseClass

from pptxlib.app import Slide
from pptxlib.core import Collection, Element
from pptxlib.shapes import Shape, Shapes
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


def test_repr_tables(tables: Tables):
    assert repr(tables) == "<Tables>"


def test_repr_table(table: Table):
    assert repr(table) == "<Table [Table]>"
