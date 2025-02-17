from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, ClassVar, Literal, overload

from win32com.client import constants

from pptxlib.colors import rgb
from pptxlib.core.base import Collection, Element
from pptxlib.core.shape import Shape

if TYPE_CHECKING:
    from collections.abc import Iterator
    from typing import Self


@dataclass(repr=False)
class Table(Shape):
    @property
    def rows(self) -> Rows:
        return Rows(self.api.Table.Rows, self)

    @property
    def columns(self) -> Columns:
        return Columns(self.api.Table.Columns, self)

    @property
    def shape(self) -> tuple[int, int]:
        return len(self.rows), len(self.columns)

    def cell(self, row: int, column: int | None = None) -> Cell:
        if column is None:
            n = len(self.columns)
            row, column = (row - 1) // n + 1, (row - 1) % n + 1

        return self.rows[row].cells[column]

    def __len__(self) -> int:
        return len(self.rows)

    @overload
    def __getitem__(self, index: int) -> Row: ...

    @overload
    def __getitem__(self, index: tuple[int, int]) -> Cell: ...

    @overload
    def __getitem__(self, index: tuple[int, slice]) -> Row: ...

    @overload
    def __getitem__(self, index: tuple[slice, int]) -> Column: ...

    def __getitem__(
        self,
        index: int | tuple[int, int] | tuple[int, slice] | tuple[slice, int],
    ) -> Cell | Row | Column:
        if isinstance(index, int):
            return self.rows[index]

        if isinstance(index, tuple):
            if isinstance(index[0], int) and isinstance(index[1], int):
                return self.cell(index[0], index[1])

            if isinstance(index[0], int) and index[1] == slice(None):
                return self.rows[index[0]]

            if index[0] == slice(None) and isinstance(index[1], int):
                return self.columns[index[1]]

        raise NotImplementedError

    def __iter__(self) -> Iterator[Row]:
        for i in range(len(self)):
            yield self[i]

    def fill(
        self,
        color: int | str | tuple[int, int, int],
        start: tuple[int, int],
        end: tuple[int, int],
    ) -> Self:
        for row in range(start[0], end[0] + 1):
            for column in range(start[1], end[1] + 1):
                cell = self[row, column]
                cell.shape.fill_color = color

        return self

    def minimize_height(self) -> None:
        for row in self.rows:
            row.height = 1


@dataclass(repr=False)
class Row(Element):
    parent: Table
    collection: Rows

    def __repr__(self) -> str:
        return f"<{self.__class__.__name__}>"

    @property
    def height(self) -> float:
        return self.api.Height

    @height.setter
    def height(self, value: float) -> None:
        self.api.Height = value

    @property
    def cells(self) -> CellRange:
        return CellRange(self.api.Cells, self)

    def __len__(self) -> int:
        return len(self.cells)

    def __getitem__(self, index: int) -> Cell:
        return self.cells[index]

    def __iter__(self) -> Iterator[Cell]:
        for i in range(len(self)):
            yield self[i]


@dataclass(repr=False)
class Rows(Collection[Row]):
    parent: Table
    type: ClassVar[type[Element]] = Row

    @property
    def height(self) -> list[float]:
        return [row.height for row in self]

    @height.setter
    def height(self, value: list[float]) -> None:
        for row, height in zip(self, value, strict=True):
            row.height = height


@dataclass(repr=False)
class Column(Element):
    parent: Table
    collection: Columns

    def __repr__(self) -> str:
        return f"<{self.__class__.__name__}>"

    @property
    def width(self) -> float:
        return self.api.Width

    @width.setter
    def width(self, value: float) -> None:
        self.api.Width = value

    @property
    def cells(self) -> CellRange:
        return CellRange(self.api.Cells, self)

    def __len__(self) -> int:
        return len(self.cells)

    def __getitem__(self, index: int) -> Cell:
        return self.cells[index]

    def __iter__(self) -> Iterator[Cell]:
        for i in range(len(self)):
            yield self[i]


@dataclass(repr=False)
class Columns(Collection[Column]):
    parent: Table
    type: ClassVar[type[Element]] = Column

    @property
    def width(self) -> list[float]:
        return [column.width for column in self]

    @width.setter
    def width(self, value: list[float]) -> None:
        for column, width in zip(self, value, strict=True):
            column.width = width


@dataclass(repr=False)
class Cell(Element):
    parent: Table
    collection: CellRange

    def __repr__(self) -> str:
        return f"<{self.__class__.__name__}>"

    @property
    def shape(self) -> Shape:
        return Shape(self.api.Shape, self.parent.parent, self.parent.collection)

    @property
    def text(self) -> str:
        return self.shape.text

    @text.setter
    def text(self, value: str) -> None:
        self.shape.text = value

    @property
    def borders(self) -> Borders:
        return Borders(self.api.Borders, self.parent)


@dataclass(repr=False)
class CellRange(Collection[Cell]):
    parent: Row | Column
    type: ClassVar[type[Element]] = Cell

    @property
    def borders(self) -> Borders:
        return Borders(self.api.Borders, self.parent.parent)


@dataclass(repr=False)
class LineFormat(Element):
    parent: Table
    collection: Borders

    def __repr__(self) -> str:
        return f"<{self.__class__.__name__}>"


@dataclass(repr=False)
class Borders(Collection[LineFormat]):
    parent: Table
    type: ClassVar[type[Element]] = LineFormat

    def __getitem__(
        self,
        index: int | Literal["bottom", "left", "right", "top"],
    ) -> LineFormat:
        if isinstance(index, int):
            return super().__getitem__(index)

        index = getattr(constants, "ppBorder" + index[0].upper() + index[1:])
        return LineFormat(self.api(index), self.parent, self)  # type: ignore


# from win32com.client import constants

# from xlviews.utils import rgb


# def set_border(table, start, end, edge_width=2, inside_width=1, edge_color=0,
#                inside_color=rgb(140, 140, 140), edge_line_style='-',
#                inside_line_style='-'):

#     if inside_width:
#         kwargs = dict(width=inside_width, color=inside_color,
#                       line_style=inside_line_style)
#         for row in range(start[0], end[0] + 1):
#             for column in range(start[1], end[1]):
#                 cell = table.Cell(row, column)
#                 set_border_cell(cell, 'right', **kwargs)
#         for column in range(start[1], end[1] + 1):
#             for row in range(start[0], end[0]):
#                 cell = table.Cell(row, column)
#                 set_border_cell(cell, 'bottom', **kwargs)

#     if edge_width:
#         kwargs = dict(width=edge_width, color=edge_color,
#                       line_style=edge_line_style)
#         for row in range(start[0], end[0] + 1):
#             cell = table.Cell(row, start[1])
#             set_border_cell(cell, 'left', **kwargs)
#             cell = table.Cell(row, end[1])
#             set_border_cell(cell, 'right', **kwargs)
#         for column in range(start[1], end[1] + 1):
#             cell = table.Cell(start[0], column)
#             set_border_cell(cell, 'top', **kwargs)
#             cell = table.Cell(end[0], column)
#             set_border_cell(cell, 'bottom', **kwargs)


# def set_font(table, start, end, size=10):
#     for row in range(start[0], end[0] + 1):
#         for column in range(start[1], end[1] + 1):
#             cell = table.cell(row, column)
#             print(cell)


# def get_borders(
#     cell: Cell | CellRange,
#     border_type: Literal["bottom","left","right","top"],
#     width: float = 1,
#     color: int | str | tuple[int, int, int] = 0,
#     line_style: Literal["-", "--"] = "-",
#     *,
#     visible: bool = True,
# ):
#     border_type_int = getattr(constants, "ppBorder" + border_type[0].upper() + border_type[1:])
#     border = cell.api.Borders(border_type_int)
#     border.Visible = visible

#     if not visible:
#         return

#     border.Weight = width
#     border.ForeColor.RGB = color

#     if line_style == "--":
#         border.DashStyle = constants.msoLineDash"
# def main():
#     import xlviews.powerpoint.table

#     table = xlviews.powerpoint.table.main()
#     set_border(table, (4, 1), (7, 2))
#     return table


# if __name__ == '__main__':
#     table = main()


# # #     def align(self, shape, pos=(0, 0)):
# # #         x, y = pos
# # #         shape.left = (2 * self.left + (self.width - shape.width) * (1 + x)) / 2
# # #         shape.top = (2 * self.top + (self.height - shape.height) * (1 - y)) / 2

# # #     def add_label(self, text, pos=(-0.98, 0.98), **kwargs):
# # #         shapes = self.parent.parent.parent.shapes
# # #         shape = shapes.add_label(text, 100, 100, **kwargs)
# # #         self.align(shape, pos=pos)

# # #     def add_picture(self, fig=None, scale=0.98, pos=(0, 0), **kwargs):
# # #         slide = self.parent.parent.parent
# # #         shape = slide.shapes.add_picture(fig=fig, width=self.width * scale, **kwargs)
# # #         self.align(shape, pos=pos)

# # #     def add_frame(self, df, pos=(0, 0), font_size=7, **kwargs):
# # #         slide = self.parent.parent.parent
# # #         shape = slide.shapes.add_frames(df, font_size=font_size, **kwargs)
# # #         self.align(shape, pos=pos)

# # #     def options(self, type):
# # #         self.type = type
# # #         return self


# # #     @property
# # #     def value(self):
# # #         values = [[cell.text for cell in row] for row in self]
# # #         if self.type:
# # #             values = self.type(values)
# # #             self.type = None
# # #         return values

# # #     def items(self):
# # #         for row in self:
# # #             for cell in row:
# # #                 yield cell.text, cell

# # #     def cells(self, row=None, column=None, start=None, dropna=True):
# # #         if row is not None or column is not None:
# # #             yield from self.cells_with_label(row, column, start, dropna)
# # #         else:
# # #             for _, cell in self.items():
# # #                 yield cell

# # #     def cells_with_label(self, row, column, start=None, dropna=True):
# # #         """
# # #         特定の行と列の値をラベルとして，セルに添付して返すジェネレータ

# # #         Parameters
# # #         ----------
# # #         row : int
# # #             ラベルに用いる行番号. 0の場合，使わない．
# # #         column : int
# # #             ラベルに用いる列番号. 0の場合，使わない．
# # #         start : tuple, optional
# # #             走査するセルの開始位置
# # #         dropna : bool, optional
# # #             ラベルがないセルをスキップするかどうか

# # #         Returns
# # #         -------
# # #         generator
# # #         """
# # #         if start is None:
# # #             start = (row + 1, column + 1)

# # #         value = self.value
# # #         column_labels = value[row - 1] if row else [None] * len(value[0])
# # #         row_labels = (
# # #             [row_value[column - 1] for row_value in value] if column else [None] * len(value)
# # #         )

# # #         for i, row in enumerate(self):
# # #             if i < start[0] - 1:
# # #                 continue
# # #             for j, cell in enumerate(row):
# # #                 if j < start[1] - 1:
# # #                     continue
# # #                 if not dropna or (row_labels[i] and column_labels[j]):
# # #                     yield cell, row_labels[i], column_labels[j]

# # #     def row(self, index):
# # #         """
# # #         特定の行のセルを返すジェネレータ

# # #         Parameters
# # #         ----------
# # #         index : int
# # #             行番号

# # #         Returns
# # #         -------
# # #         generator
# # #         """
# # #         for column in range(self.shape[1]):
# # #             yield self.cell(index, column + 1)

# # #     def column(self, index):
# # #         """
# # #         特定の列のセルを返すジェネレータ

# # #         Parameters
# # #         ----------
# # #         index : int
# # #             列番号

# # #         Returns
# # #         -------
# # #         generator
# # #         """
# # #         for row in range(self.shape[0]):
# # #             yield self.cell(row + 1, index)

# # #     def apply(self, func, *args, pattern=None, **kwargs):
# # #         """
# # #         テーブルの各セルに対して関数を適用する．

# # #         Parameters
# # #         ----------
# # #         func
# # #         args
# # #         pattern
# # #         kwargs
# # #         """
# # #         if isinstance(pattern, str):
# # #             match = re.compile(pattern).match
# # #         else:
# # #             match = None

# # #         for cell in self.cells():
# # #             if match and not match(cell.value):
# # #                 continue
# # #             func(cell, *args, **kwargs)

# # #     def clean(self):
# # #         """全角空白のみのセルを空文字に変更する．"""
# # #         rows, columns = self.shape
# # #         for row in range(rows):
# # #             for column in range(columns):
# # #                 cell = self.cell(row + 1, column + 1)
# # #                 if cell.text == "\u3000":
# # #                     cell.text = ""
