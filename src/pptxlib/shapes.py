from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, ClassVar, Literal

from win32com.client import constants

from pptxlib.base import Collection, Element
from pptxlib.utils import rgb

if TYPE_CHECKING:
    from win32com.client import DispatchBaseClass

    from pptxlib.core import Slide
    from pptxlib.tables import Cell


@dataclass(repr=False)
class Shape(Element):
    parent: Slide | Cell

    @classmethod
    def get_parent(cls, collection: Shapes) -> Slide:
        return collection.parent

    @property
    def text_range(self) -> DispatchBaseClass:
        return self.api.TextFrame.TextRange

    @property
    def text(self) -> str:
        return self.text_range.Text

    @text.setter
    def text(self, text: str) -> None:
        self.text_range.Text = text

    @property
    def slide(self) -> Slide:
        from pptxlib.core import Slide

        if isinstance(self.parent, Slide):
            return self.parent

        raise NotImplementedError

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
        slide = self.slide

        if value == "center":
            value = (slide.width - self.width) / 2
        elif value < 0:
            value = slide.width - self.width + value

        self.api.Left = value
        return value

    @top.setter
    def top(self, value: float | Literal["center"]) -> float:
        slide = self.slide

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

    @property
    def font(self) -> DispatchBaseClass:
        return self.text_range.Font

    @property
    def font_name(self) -> str:
        return self.font.Name

    @font_name.setter
    def font_name(self, name: str) -> None:
        self.font.Name = name

    @property
    def font_size(self) -> float:
        return self.font.Size

    @font_size.setter
    def font_size(self, size: float):
        self.font.Size = size

    @property
    def bold(self) -> bool:
        return self.font.Bold == constants.msoTrue

    @bold.setter
    def bold(self, value: bool):
        self.font.Bold = value

    @property
    def italic(self) -> bool:
        return self.font.Italic == constants.msoTrue

    @italic.setter
    def italic(self, value: bool):
        self.font.Italic = value

    @property
    def color(self):
        return self.font.Color.RGB

    @color.setter
    def color(self, value: int | str | tuple[int, int, int]):
        self.font.Color.RGB = rgb(value)

    @property
    def fill_color(self):
        return self.api.Fill.ForeColor.RGB

    @fill_color.setter
    def fill_color(self, value):
        self.api.Fill.ForeColor.RGB = rgb(value)

    @property
    def line_color(self):
        return self.api.Line.ForeColor.RGB

    @line_color.setter
    def line_color(self, value):
        self.api.Line.ForeColor.RGB = rgb(value)

    @property
    def line_weight(self):
        return self.api.Line.Weight

    @line_weight.setter
    def line_weight(self, value):
        self.api.Line.Weight = value

    def set_style(
        self,
        font=None,
        size=None,
        bold=None,
        italic=None,
        color=None,
        fill_color=None,
        line_weight=None,
        line_color=None,
    ):
        if font is not None:
            self.font_name = font
        if size is not None:
            self.font_size = size
        if bold is not None:
            self.bold = bold
        if italic is not None:
            self.italic = italic
        if color is not None:
            self.color = color
        if fill_color is not None:
            self.fill_color = fill_color
        if line_weight is not None:
            self.line_weight = line_weight
        if line_color is not None:
            self.line_color = line_color


@dataclass(repr=False)
class Shapes(Collection[Shape]):
    parent: Slide
    type: ClassVar[type[Element]] = Shape

    @property
    def title(self) -> Shape:
        return Shape(self.api.Title, self.parent)

    def add(
        self,
        kind: int | str,
        left: float,
        top: float,
        width: float,
        height: float,
        text: str = "",
        **kwargs,
    ) -> Shape:
        if isinstance(kind, str):
            kind = getattr(constants, f"msoShape{kind}")

        api = self.api.AddShape(kind, left, top, width, height)
        shape = Shape(api, self.parent)
        shape.text = text
        shape.set_style(**kwargs)

        return shape

    def add_label(
        self,
        text: str,
        left: float,
        top: float,
        width: float = 72,
        height: float = 72,
        *,
        auto_size: bool = True,
        **kwargs,
    ) -> Shape:
        orientation = constants.msoTextOrientationHorizontal
        api = self.api.AddLabel(orientation, left, top, width, height)

        if auto_size is False:
            api.TextFrame.AutoSize = False

        shape = Shape(api, self.parent)
        shape.text = text
        shape.set_style(**kwargs)

        return shape

    def add_table(
        self,
        num_rows: int,
        num_columns: int,
        left: float = 100,
        top: float = 100,
        width: float = 100,
        height: float = 100,
    ) -> Shape:
        api = self.api.AddTable(num_rows, num_columns, left, top, width, height)
        return Shape(api, self.parent)


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
