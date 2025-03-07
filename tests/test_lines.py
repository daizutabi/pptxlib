from pptxlib.tables import Cell, CellRange


def test_cell_line_format(cell: Cell):
    line_format = cell.borders("left")
    assert line_format.api.__class__.__name__ == "LineFormat"
    assert line_format.__class__.__name__ == "LineFormat"


def test_cell_line_format_parent(cell: Cell):
    line_format = cell.borders("left")
    assert line_format.api.Parent.__class__.__name__ == "Borders"
    assert line_format.parent.__class__.__name__ == "Borders"


def test_cell_range_line_format(cell_range: CellRange):
    line_format = cell_range.borders("left")
    assert line_format.api.__class__.__name__ == "LineFormat"
    assert line_format.__class__.__name__ == "LineFormat"


def test_cell_range_line_format_parent(cell_range: CellRange):
    line_format = cell_range.borders("left")
    assert line_format.api.Parent.__class__.__name__ == "Borders"
    assert line_format.parent.__class__.__name__ == "Borders"
