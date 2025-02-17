from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, ClassVar, Literal

from win32com.client import constants

from pptxlib.core.base import Collection, Element
from pptxlib.core.shape import Shape
from pptxlib.line import LineFormat

if TYPE_CHECKING:
    from collections.abc import Iterator

    from .slide import Slide


@dataclass(repr=False)
class Table(Element):
    parent: Shape
    collection: Tables

    def delete(self) -> None:
        self.parent.delete()

    @property
    def left(self) -> float:
        return self.parent.left

    @left.setter
    def left(self, value: float) -> None:
        self.parent.left = value

    @property
    def top(self) -> float:
        return self.parent.top

    @top.setter
    def top(self, value: float) -> None:
        self.parent.top = value

    @property
    def width(self) -> float:
        return self.parent.width

    @width.setter
    def width(self, value: float) -> None:
        self.parent.width = value

    @property
    def height(self) -> float:
        return self.parent.height

    @height.setter
    def height(self, value: float) -> None:
        self.parent.height = value

    def minimize_height(self) -> None:
        for row in self.rows:
            row.height = 1

    @property
    def rows(self) -> Rows:
        return Rows(self.api.Rows, self)

    @property
    def columns(self) -> Columns:
        return Columns(self.api.Columns, self)

    @property
    def shape(self) -> tuple[int, int]:
        return len(self.rows), len(self.columns)

    def cell(self, row: int, column: int | None = None) -> Cell:
        if column is None:
            n = len(self.columns)
            row, column = (row - 1) // n + 1, (row - 1) % n + 1

        return Cell(self.api.Cell(row, column), self)


# def clear(table):
#     table.FirstCol = False
#     table.FirstRow = False
#     table.HorizBanding = False

#     nrows = len(table.Columns(1).Cells)
#     ncols = len(table.Rows(1).Cells)
#     for row, column in product(range(nrows), range(ncols)):
#         cell = table.Cell(row + 1, column + 1)
#         if row == 0:
#             set_border_cell(cell, 'top', visible=False)
#         if column == 0:
#             set_border_cell(cell, 'left', visible=False)
#         set_border_cell(cell, 'right', visible=False)
#         set_border_cell(cell, 'bottom', visible=False)
#         cell.Shape.Fill.Visible = False


@dataclass(repr=False)
class Tables(Collection[Table]):
    parent: Slide
    type: ClassVar[type[Element]] = Table

    def __iter__(self) -> Iterator[Table]:
        for index in range(self.api.Count):
            api = self.api(index + 1)  # type: ignore

            if api.HasTable:
                yield Table(api.Table, Shape(api, self.parent))

    def __len__(self) -> int:
        return len(list(iter(self)))

    def __call__(self, index: int | None = None) -> Table:
        tables = list(iter(self))
        if index is None:
            index = len(tables)

        return tables[index - 1]

    def add(
        self,
        num_rows: int,
        num_columns: int,
        left: float = 100,
        top: float = 100,
        width: float = 100,
        height: float = 100,
    ) -> Table:
        api = self.api.AddTable(num_rows, num_columns, left, top, width, height)
        return Table(api.Table, Shape(api, self.parent))
