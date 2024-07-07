from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, ClassVar

from pptxlib.core import Collection, Element
from pptxlib.shapes import Shape

if TYPE_CHECKING:
    from collections.abc import Iterator

    from pptxlib.app import Slide


@dataclass(repr=False)
class Table(Element):
    parent: Shape

    def delete(self) -> None:
        self.parent.delete()

    @property
    def left(self) -> float:
        return self.parent.left

    @left.setter
    def left(self, value) -> None:
        self.parent.left = value

    @property
    def top(self) -> float:
        return self.parent.top

    @top.setter
    def top(self, value) -> None:
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
        return Rows(self)

    @property
    def columns(self) -> Columns:
        return Columns(self)

    @property
    def shape(self) -> tuple[int, int]:
        return len(self.rows), len(self.columns)

    def cell(self, row: int, column: int | None = None) -> Cell:
        if column is None:
            n = len(self.columns)
            row, column = (row - 1) // n + 1, (row - 1) % n + 1

        return Cell(self.api.Cell(row, column), self)


@dataclass(repr=False)
class Tables(Collection[Table]):
    parent: Slide
    type: ClassVar[type[Element]] = Table

    def __post_init__(self):
        self.api = self.parent.api.Shapes
        self.app = self.parent.app

    def __iter__(self) -> Iterator[Table]:
        for index in range(self.api.Count):
            api = self.api(index + 1)  # type: ignore

            if api.HasTable:
                yield Table(api.Table, Shape(api, self.parent.shapes))

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
        return Table(api.Table, Shape(api, self.parent.shapes))


@dataclass(repr=False)
class Row(Element):
    parent: Rows

    @property
    def height(self) -> float:
        return self.api.Height

    @height.setter
    def height(self, value: float) -> None:
        self.api.Height = value


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
    parent: Columns

    @property
    def width(self) -> float:
        return self.api.Width

    @width.setter
    def width(self, value: float) -> None:
        self.api.Width = value


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

    @property
    def shape(self):
        return Shape(self.api.Shape, parent=self)

    @property
    def left(self):
        return self.shape.left

    @property
    def top(self):
        return self.shape.top

    @property
    def width(self):
        return self.shape.width

    @property
    def height(self):
        return self.shape.height

    @property
    def text(self):
        return self.shape.text

    @text.setter
    def text(self, value):
        self.shape.text = value

    @property
    def value(self):
        return self.text

    @value.setter
    def value(self, value):
        self.text = value


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
