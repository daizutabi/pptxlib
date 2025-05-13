from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, ClassVar, Literal

from win32com.client import constants

from pptxlib.colors import rgb
from pptxlib.core.base import Base, Collection, Element
from pptxlib.core.font import Font

if TYPE_CHECKING:
    from typing import Self

    from win32com.client import DispatchBaseClass

    from .slide import Slide
    from .table import Table


@dataclass(repr=False)
class Color(Base):
    @property
    def color(self) -> int:
        return self.api.ForeColor.RGB

    @color.setter
    def color(self, value: int | str | tuple[int, int, int]) -> None:
        self.api.ForeColor.RGB = rgb(value)

    @property
    def alpha(self) -> float:
        return self.api.Transparency

    @alpha.setter
    def alpha(self, value: float) -> None:
        self.api.Transparency = value

    def set(
        self,
        color: int | str | tuple[int, int, int] | None = None,
        alpha: float | None = None,
    ) -> Self:
        if color is not None:
            self.color = color
        if alpha is not None:
            self.alpha = alpha

        return self

    def update(self, color: Color) -> None:
        self.color = color.color
        self.alpha = color.alpha


@dataclass(repr=False)
class Fill(Color):
    pass


@dataclass(repr=False)
class Line(Color):
    @property
    def weight(self) -> float:
        return self.api.Weight

    @weight.setter
    def weight(self, value: float) -> None:
        self.api.Weight = value

    def set(
        self,
        color: int | str | tuple[int, int, int] | None = None,
        alpha: float | None = None,
        weight: float | None = None,
    ) -> Self:
        if color is not None:
            self.color = color
        if alpha is not None:
            self.alpha = alpha
        if weight is not None:
            self.weight = weight

        return self

    def update(self, line: Line) -> None:
        self.color = line.color
        self.alpha = line.alpha
        self.weight = line.weight


@dataclass(repr=False)
class Shape(Element):
    parent: Slide
    collection: Shapes

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
        slide = self.parent

        if value == "center":
            value = (slide.width - self.width) / 2
        elif value < 0:
            value = slide.width - self.width + value

        self.api.Left = value
        return value

    @top.setter
    def top(self, value: float | Literal["center"]) -> float:
        slide = self.parent

        if value == "center":
            value = (slide.height - self.height) / 2
        elif value < 0:
            value = slide.height - self.height + value

        self.api.Top = value
        return value

    @width.setter
    def width(self, value: float) -> None:
        self.api.Width = value

    @height.setter
    def height(self, value: float) -> None:
        self.api.Height = value

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
    def font(self) -> Font:
        return Font(self.text_range.Font)

    @font.setter
    def font(self, value: Font) -> None:
        self.font.update(value)

    @property
    def fill(self) -> Fill:
        return Fill(self.api.Fill)

    @fill.setter
    def fill(self, value: Fill) -> None:
        self.fill.update(value)

    @property
    def line(self) -> Line:
        return Line(self.api.Line)

    @line.setter
    def line(self, value: Line) -> None:
        self.line.update(value)


@dataclass(repr=False)
class Shapes(Collection[Shape]):
    parent: Slide
    type: ClassVar[type[Element]] = Shape

    @property
    def title(self) -> Shape:
        return Shape(self.api.Title, self.parent, self)

    def add(
        self,
        kind: int | str,
        left: float,
        top: float,
        width: float,
        height: float,
        text: str = "",
    ) -> Shape:
        if isinstance(kind, str):
            kind = getattr(constants, f"msoShape{kind}")

        api = self.api.AddShape(kind, left, top, width, height)
        shape = Shape(api, self.parent, self)
        shape.text = text

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
    ) -> Shape:
        orientation = constants.msoTextOrientationHorizontal
        api = self.api.AddLabel(orientation, left, top, width, height)

        if auto_size is False:
            api.TextFrame.AutoSize = False

        label = Shape(api, self.parent, self)
        label.text = text
        return label

    def add_table(
        self,
        num_rows: int,
        num_columns: int,
        left: float = 100,
        top: float = 100,
        width: float = 100,
        height: float = 100,
    ) -> Table:
        from .table import Table

        api = self.api.AddTable(num_rows, num_columns, left, top, width, height)
        return Table(api, self.parent, self)


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
