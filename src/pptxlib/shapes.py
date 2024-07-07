from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, ClassVar, Literal

from win32com.client import constants

from pptxlib.core import Collection, Element

if TYPE_CHECKING:
    from win32com.client import DispatchBaseClass

    from pptxlib.app import Slide


class Shape(Element):
    parent: Shapes

    @property
    def text_range(self) -> DispatchBaseClass:
        return self.api.TextFrame.TextRange

    @property
    def text(self) -> str:
        try:
            return self.text_range.Text
        except AttributeError:
            return ""

    @text.setter
    def text(self, text: str) -> None:
        self.text_range.Text = text

    @property
    def slide(self) -> Slide:
        return self.parent.parent

    @property
    def left(self) -> float:
        return self.api.Left

    @property
    def top(self) -> float:
        return self.api.Top

    @property
    def width(self) -> float:
        return self.api.Width

    @property
    def height(self) -> float:
        return self.api.Height

    @left.setter
    def left(self, value: float | Literal["center"]) -> float:
        slide = self.parent.parent

        if value == "center":
            value = (slide.width - self.width) / 2
        elif value < 0:
            value = slide.width - self.width + value

        self.api.Left = value
        return value

    @top.setter
    def top(self, value: float | Literal["center"]) -> float:
        slide = self.parent.parent

        if value == "center":
            value = (slide.height - self.height) / 2
        elif value < 0:
            value = slide.height - self.height + value

        self.api.Top = value
        return value

    @width.setter
    def width(self, value):
        self.api.Width = value

    @height.setter
    def height(self, value):
        self.api.Height = value


#     @property
#     def table(self):
#         if self.api.HasTable:
#             return Table(self.api.Table, parent=self)


#     @property
#     def font(self):
#         return self.text_range.Font

#     @property
#     def size(self):
#         return self.font.Size

#     @size.setter
#     def size(self, value):
#         self.font.Size = value

#     @property
#     def bold(self):
#         return self.font.Bold

#     @bold.setter
#     def bold(self, value):
#         self.font.Bold = value

#     @property
#     def italic(self):
#         return self.font.Italic

#     @italic.setter
#     def italic(self, value):
#         self.font.Italic = value

#     @property
#     def color(self):
#         return self.font.Color.RGB

#     @color.setter
#     def color(self, value):
#         self.font.Color.RGB = rgb(value)

#     @property
#     def fill_color(self):
#         return self.api.Fill.ForeColor.RGB

#     @fill_color.setter
#     def fill_color(self, value):
#         self.api.Fill.ForeColor.RGB = rgb(value)

#     @property
#     def line_color(self):
#         return self.api.Line.ForeColor.RGB

#     @line_color.setter
#     def line_color(self, value):
#         self.api.Line.ForeColor.RGB = rgb(value)

#     @property
#     def line_weight(self):
#         return self.api.Line.Weight

#     @line_weight.setter
#     def line_weight(self, value):
#         self.api.Line.Weight = value

#     def set_style(
#         self,
#         size=None,
#         bold=None,
#         italic=None,
#         color=None,
#         fill_color=None,
#         line_color=None,
#         line_weight=None,
#     ):
#         if size is not None:
#             self.size = size
#         if bold is not None:
#             self.bold = bold
#         if italic is not None:
#             self.italic = italic
#         if color is not None:
#             self.color = color
#         if fill_color is not None:
#             self.fill_color = fill_color
#         if line_color is not None:
#             self.line_color = line_color
#         if line_weight is not None:
#             self.line_weight = line_weight


@dataclass(repr=False)
class Shapes(Collection[Shape]):
    parent: Slide
    type: ClassVar[type[Element]] = Shape

    @property
    def title(self) -> Shape:
        return Shape(self.api.Title, self)


#     def __init__(self, parent):
#         super().__init__(parent, Shape)

#     def add_picture(self, path=None, left=0, top=0, width=None, height=None, scale=1, **kwargs):
#         if not isinstance(path, str):
#             return self.add_picture_matplotlib(path, left, top, width, scale, **kwargs)

#         path = os.path.abspath(path)

#         # オリジナルの大きさするために以下のwidth, heightの指定が必要．
#         # width = 300 を width = 100 等とすると小さく表示される．なぜ？
#         if width is None:
#             width_ = 300
#             scale_ = scale
#         else:
#             width_ = width
#             scale_ = None
#         with PIL.Image.open(path) as image:
#             size = image.size
#         if height is None:
#             height_ = width_ * size[1] / size[0]
#         else:
#             height_ = height

#         shape = self.api.AddPicture(
#             FileName=path,
#             LinkToFile=False,
#             SaveWithDocument=True,
#             Left=left,
#             Top=top,
#             Width=width_,
#             Height=height_,
#         )

#         if scale_:
#             shape.ScaleWidth(scale, 1)
#             shape.ScaleHeight(scale, 1)
#         return Shape(shape, parent=self.parent)

#     def add_picture_matplotlib(self, fig=None, left=0, top=0, width=None, scale=1, dpi=None):
#         if fig is None:
#             fig = plt.gcf()
#         elif not hasattr(fig, "savefig"):
#             fig = fig.figure

#         with tempfile.TemporaryDirectory() as directory:
#             path = os.path.join(directory, "a.png")
#             fig.savefig(path, dpi=dpi, bbox_inches="tight")
#             return self.add_picture(path, left, top, width, None, scale)

#     def add_frame(self, df, left=None, top=None, **kwargs):
#         excel = xw.App(visible=False)
#         book = excel.books(1)
#         sheet = book.sheets(1)
#         sf = SheetFrame(sheet, 2, 2, data=df, **kwargs)
#         sf.range().api.Copy()
#         self.parent.api.Select()
#         n = len(self)
#         api = self.parent.parent.parent.api
#         api.CommandBars.ExecuteMso("PasteSourceFormatting")
#         while len(self) == n:  # wait for creating a table
#             pass
#         book.api.CutCopyMode = False  # Don't show confirm message
#         excel.quit()
#         excel.kill()

#         shape = Shape.from_collection(self)  # type: Shape
#         if left:
#             shape.left = left
#         if top:
#             shape.top = top

#         shape.table.clean()
#         shape.table.minimize_height()

#         return shape

#     def add_range(self, range_, data_type=2, left=None, top=None, width=None, height=None):
#         """

#         Parameters
#         ----------
#         range_ : xlwings.Range
#         data_type : int
#             0: ppPasteDefault (既定の内容)
#             1: ppPasteBitmap (ビットマップ)
#             2: ppPasteEnhancedMetafile (拡張メタファイル)
#             4: ppPasteGIF
#             8: ppPasteHTML
#             5: ppPasteJPG
#             3: ppPasteMetafilePicture
#             10: ppPasteOLEObject
#             6: ppPastePNG
#             9: ppPasteRTF
#             11: ppPasteShape
#             7: ppPasteText
#         left, top, width, height: int, optional
#             図形のディメンジョン

#         Returns
#         -------
#         Shape
#         """
#         range_.api.Copy()
#         shape = self.api.PasteSpecial(data_type)
#         # range.sheet.book.api.CutCopyMode = False
#         shape.LockAspectRatio = 0
#         if left:
#             shape.Left = left
#         if top:
#             shape.Top = top
#         if width:
#             shape.Width = width
#         if height:
#             shape.Height = height

#         return Shape(shape, parent=self.parent)

#     def add_chart(self, chart, left=None, top=None, width=None, height=None, scale=None):
#         """
#         Parameters
#         ----------
#         chart : xlwings or altairのチャート
#         left, top, width, height: int, optional
#             図形のディメンジョン

#         Returns
#         -------
#         Shape
#         """
#         if isinstance(chart, list):
#             charts = chart
#             left_ = charts[0].left
#             top_ = charts[0].top
#             shapes = []
#             for chart in charts:
#                 shape = self.add_chart(
#                     chart,
#                     left=chart.left - left_ + left,
#                     top=chart.top - top_ + top,
#                     width=width,
#                     height=height,
#                 )
#                 shapes.append(shape)
#             return shapes

#         if hasattr(chart, "save"):
#             return self.add_chart_altair(chart, left, top, width, height, scale)
#         else:
#             return self.add_chart_xlwings(chart, left, top, width, height)

#     def add_chart_altair(self, chart, left=None, top=None, width=None, height=None, scale=None):
#         """
#         Parameters
#         ----------
#         chart : altairのチャート
#         left, top, width, height: int, optional
#             図形のディメンジョン

#         Returns
#         -------
#         Shape
#         """
#         with tempfile.TemporaryDirectory() as directory:
#             path = os.path.join(directory, "a.png")
#             chart.save(path)
#             return self.add_picture(path, left, top, width, height, scale)

#     def add_chart_xlwings(self, chart, left=None, top=None, width=None, height=None):
#         """
#         Parameters
#         ----------
#         chart : xlwingsのチャート
#         left, top, width, height: int, optional
#             図形のディメンジョン

#         Returns
#         -------
#         Shape
#         """
#         chart.api[0].Copy()
#         shape = self.api.Paste()
#         if left:
#             shape.Left = left
#         if top:
#             shape.Top = top
#         if width:
#             shape.Width = width
#         if height:
#             shape.Height = height

#         return Shape(shape, parent=self.parent)

#     def add_label(self, text, left, top, width=72, height=72, auto_size=True, **kwargs):
#         orientation = 1  # msoTextOrientationHorizontal
#         shape = self.api.AddLabel(orientation, left, top, width, height)
#         if auto_size is False:
#             shape.TextFrame.AutoSize = False
#         shape = Shape(shape, parent=self.parent)
#         shape.text = text
#         shape.set_style(**kwargs)
#         return shape

#     def add_shape(self, type, left, top, width, height, text=None, **kwargs):
#         shape = self.api.AddShape(type, left, top, width, height)
#         shape = Shape(shape, parent=self.parent)
#         if text:
#             shape.text = text
#         shape.set_style(**kwargs)
#         return shape

#     def add_table(self, df, left=None, top=None, width=None, height=None, merge=True, **kwargs):
#         shape = create_table(
#             self, df, left=100, top=100, width=300, height=300, merge=merge, **kwargs
#         )
#         table = shape.table
#         if width:
#             table.width = width
#         if height:
#             table.height = height
#         if left is not None:
#             shape.left = left
#         if top is not None:
#             shape.top = top

#         return table


# # class Tables(CollectionBase):
# #     def __init__(self, slide):
# #         tables = [shape.table for shape in slide.shapes]
# #         self._tables = [table for table in tables if table is not None]

# #     def __getitem__(self, item):
# #         return self._tables[item]

# #     def __len__(self):
# #         return len(self._tables)

# #     def __iter__(self):
# #         return iter(self._tables)

# #     def __call__(self, index):
# #         return self[index - 1]


# # class Table(Element):
# #     def __init__(self, *args, **kwargs):
# #         super().__init__(*args, **kwargs)
# #         self.type = None

# #     def __len__(self):
# #         return len(self.rows)

# #     def __iter__(self):
# #         rows, columns = self.shape

# #         def row_iter(row):
# #             for column in range(columns):
# #                 yield self.cell(row + 1, column + 1)

# #         for row in range(rows):
# #             yield row_iter(row)

# #     def __call__(self, i, j):
# #         return self.cell(i, j)

# #     def __repr__(self):
# #         parent = repr(self.parent)
# #         parent = parent[parent.index(" ") + 1 : -1]
# #         return f"<Table {parent}{self.shape}>"

# #     def cell(self, i, j=None):
# #         if j is None:
# #             n = len(self.columns)
# #             i, j = (i - 1) // n + 1, (i - 1) % n + 1
# #         return Cell(self.api.Cell(i, j), parent=self)

# #     def options(self, type):
# #         self.type = type
# #         return self

# #     @property
# #     def shape(self):
# #         return len(self.rows), len(self.columns)

# #     @property
# #     def left(self):
# #         return self.parent.left

# #     @left.setter
# #     def left(self, value):
# #         self.parent.left = value

# #     @property
# #     def top(self):
# #         return self.parent.top

# #     @top.setter
# #     def top(self, value):
# #         self.parent.top = value

# #     @property
# #     def value(self):
# #         values = [[cell.text for cell in row] for row in self]
# #         if self.type:
# #             values = self.type(values)
# #             self.type = None
# #         return values

# #     def items(self):
# #         for row in self:
# #             for cell in row:
# #                 yield cell.text, cell

# #     def cells(self, row=None, column=None, start=None, dropna=True):
# #         if row is not None or column is not None:
# #             yield from self.cells_with_label(row, column, start, dropna)
# #         else:
# #             for _, cell in self.items():
# #                 yield cell

# #     def cells_with_label(self, row, column, start=None, dropna=True):
# #         """
# #         特定の行と列の値をラベルとして，セルに添付して返すジェネレータ

# #         Parameters
# #         ----------
# #         row : int
# #             ラベルに用いる行番号. 0の場合，使わない．
# #         column : int
# #             ラベルに用いる列番号. 0の場合，使わない．
# #         start : tuple, optional
# #             走査するセルの開始位置
# #         dropna : bool, optional
# #             ラベルがないセルをスキップするかどうか

# #         Returns
# #         -------
# #         generator
# #         """
# #         if start is None:
# #             start = (row + 1, column + 1)

# #         value = self.value
# #         column_labels = value[row - 1] if row else [None] * len(value[0])
# #         row_labels = (
# #             [row_value[column - 1] for row_value in value] if column else [None] * len(value)
# #         )

# #         for i, row in enumerate(self):
# #             if i < start[0] - 1:
# #                 continue
# #             for j, cell in enumerate(row):
# #                 if j < start[1] - 1:
# #                     continue
# #                 if not dropna or (row_labels[i] and column_labels[j]):
# #                     yield cell, row_labels[i], column_labels[j]

# #     def row(self, index):
# #         """
# #         特定の行のセルを返すジェネレータ

# #         Parameters
# #         ----------
# #         index : int
# #             行番号

# #         Returns
# #         -------
# #         generator
# #         """
# #         for column in range(self.shape[1]):
# #             yield self.cell(index, column + 1)

# #     def column(self, index):
# #         """
# #         特定の列のセルを返すジェネレータ

# #         Parameters
# #         ----------
# #         index : int
# #             列番号

# #         Returns
# #         -------
# #         generator
# #         """
# #         for row in range(self.shape[0]):
# #             yield self.cell(row + 1, index)

# #     def apply(self, func, *args, pattern=None, **kwargs):
# #         """
# #         テーブルの各セルに対して関数を適用する．

# #         Parameters
# #         ----------
# #         func
# #         args
# #         pattern
# #         kwargs
# #         """
# #         if isinstance(pattern, str):
# #             match = re.compile(pattern).match
# #         else:
# #             match = None

# #         for cell in self.cells():
# #             if match and not match(cell.value):
# #                 continue
# #             func(cell, *args, **kwargs)

# #     def clean(self):
# #         """全角空白のみのセルを空文字に変更する．"""
# #         rows, columns = self.shape
# #         for row in range(rows):
# #             for column in range(columns):
# #                 cell = self.cell(row + 1, column + 1)
# #                 if cell.text == "\u3000":
# #                     cell.text = ""

# #     def minimize_height(self):
# #         for row in self.rows:
# #             row.height = 1

# #     @property
# #     def columns(self):
# #         return Columns(self)

# #     @property
# #     def rows(self):
# #         return Rows(self)

# #     @property
# #     def width(self):
# #         return [column.width for column in self.columns]

# #     @width.setter
# #     def width(self, value):
# #         if isinstance(value, list):
# #             for column, width in zip(self.columns, value, strict=False):
# #                 column.width = width
# #         else:
# #             self.parent.width = value

# #     @property
# #     def height(self):
# #         return [row.height for row in self.rows]

# #     @height.setter
# #     def height(self, value):
# #         if isinstance(value, list):
# #             for row, height in zip(self.rows, value, strict=False):
# #                 row.height = value
# #         else:
# #             self.parent.height = value


# # class Columns(Collection):
# #     def __init__(self, parent):
# #         super().__init__(parent, Column)


# # class Column(Element):
# #     @property
# #     def width(self):
# #         return self.api.Width

# #     @width.setter
# #     def width(self, value):
# #         self.api.Width = value


# # class Rows(Collection):
# #     def __init__(self, parent):
# #         super().__init__(parent, Row)


# # class Row(Element):
# #     @property
# #     def height(self):
# #         return self.api.Height

# #     @height.setter
# #     def height(self, value):
# #         self.api.Height = value


# # class Cell(Element):
# #     def __repr__(self):
# #         parent = repr(self.parent.parent)
# #         parent = parent[parent.index(" ") + 1 : -1]
# #         return f"<Cell {parent}>"

# #     @property
# #     def shape(self):
# #         return Shape(self.api.Shape, parent=self)

# #     @property
# #     def text(self):
# #         return self.shape.text

# #     @text.setter
# #     def text(self, value):
# #         self.shape.text = value

# #     @property
# #     def value(self):
# #         return self.text

# #     @value.setter
# #     def value(self, value):
# #         self.text = value

# #     @property
# #     def left(self):
# #         return self.shape.left

# #     @property
# #     def top(self):
# #         return self.shape.top

# #     @property
# #     def width(self):
# #         return self.shape.width

# #     @property
# #     def height(self):
# #         return self.shape.height

# #     def align(self, shape, pos=(0, 0)):
# #         x, y = pos
# #         shape.left = (2 * self.left + (self.width - shape.width) * (1 + x)) / 2
# #         shape.top = (2 * self.top + (self.height - shape.height) * (1 - y)) / 2

# #     def add_label(self, text, pos=(-0.98, 0.98), **kwargs):
# #         shapes = self.parent.parent.parent.shapes
# #         shape = shapes.add_label(text, 100, 100, **kwargs)
# #         self.align(shape, pos=pos)

# #     def add_picture(self, fig=None, scale=0.98, pos=(0, 0), **kwargs):
# #         slide = self.parent.parent.parent
# #         shape = slide.shapes.add_picture(fig=fig, width=self.width * scale, **kwargs)
# #         self.align(shape, pos=pos)

# #     def add_frame(self, df, pos=(0, 0), font_size=7, **kwargs):
# #         slide = self.parent.parent.parent
# #         shape = slide.shapes.add_frames(df, font_size=font_size, **kwargs)
# #         self.align(shape, pos=pos)


# # def _prepare_frame():
# #     import pandas as pd

# #     df = pd.DataFrame([[1, 2, 3], [4, 5, 6], [7, 8, 9], [10, 11, 12]])
# #     df.index = [["a", "a", "a", "a"], ["x", "y", "y", "z"]]
# #     df.columns = [["A", "B", "B"], ["s", "t", "t"]]
# #     df.index.names = ["i1", "i2"]
# #     df.columns.names = ["c1", "c2"]
# #     return df


# def main():
#     import xlviews as xv

#     pp = xv.PowerPoint()

#     import altair as alt

#     # load a simple dataset as a pandas DataFrame
#     from vega_datasets import data

#     cars = data.cars()

#     chart = (
#         alt.Chart(cars)
#         .mark_point()
#         .encode(x="Horsepower", y="Miles_per_Gallon", color="Origin")
#         .interactive()
#     )

#     chart.save("a.svg")
#     from PIL import Image

#     img = Image.open("a.png")
#     img.save("aa.png", dpi=(200, 200))

#     pp.add_chart(chart, 100, 100, scale=1)

#     # df = _prepare_frame()
#     #
#     # for table in pp.tables:
#     #     table.parent.api.Delete()
#     #
#     # pp.add_table(df, left=200, top=100, width=300, height=200,
#     #              columns_name=True, index_name=False)

#     import matplotlib.pyplot as plt

#     fig = plt.subplot(111)
#     fig.plot([1, 2])
#     fig.figure.savefig("bb.png", dpi=200)
#     pp.add_picture(fig, left=100, top=10, dpi=100)


# if __name__ == "__main__":
#     main()
