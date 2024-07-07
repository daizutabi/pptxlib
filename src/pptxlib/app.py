from __future__ import annotations

from dataclasses import dataclass
from typing import ClassVar

import win32com.client
from win32com.client import constants

from pptxlib.client import ensure_modules
from pptxlib.core import Base, Collection, Element
from pptxlib.shapes import Shapes
from pptxlib.tables import Tables


@dataclass(repr=False)
class PowerPoint(Base):
    def __post_init__(self):
        ensure_modules()
        self.api = win32com.client.Dispatch("PowerPoint.Application")  # type: ignore
        self.app = self.api

    @property
    def presentations(self):
        return Presentations(self)

    def quit(self):
        self.api.Quit()


@dataclass(repr=False)
class Presentation(Element):
    parent: Presentations

    def close(self):
        self.api.Close()

    @property
    def slides(self):
        return Slides(self)

    @property
    def width(self) -> float:
        return self.api.PageSetup.SlideWidth

    @property
    def height(self) -> float:
        return self.api.PageSetup.SlideHeight


@dataclass(repr=False)
class Presentations(Collection[Presentation]):
    parent: PowerPoint
    type: ClassVar[type[Element]] = Presentation

    def add(self) -> Presentation:
        api = self.api.Add()
        return Presentation(api, self)

    @property
    def active(self) -> Presentation:
        api = self.app.ActivePresentation
        return Presentation(api, self)

    # def open(self, filename):
    #     filename = os.path.abspath(filename)
    #     prs = self.api.Open(filename)
    #     return Presentation(prs, parent=self.parent)


@dataclass(repr=False)
class Slide(Element):
    parent: Slides

    @property
    def shapes(self) -> Shapes:
        return Shapes(self)

    @property
    def tables(self) -> Tables:
        return Tables(self)

    @property
    def title(self) -> str:
        return self.shapes.title.text if len(self.shapes) else ""

    @title.setter
    def title(self, text):
        if len(self.shapes):
            self.shapes.title.text = text

    @property
    def width(self) -> float:
        return self.parent.parent.width

    @property
    def height(self) -> float:
        return self.parent.parent.height


class Slides(Collection[Slide]):
    parent: Presentation
    type: ClassVar[type[Element]] = Slide

    def add(self, index: int | None = None, layout=None):
        if index is None:
            index = len(self) + 1

        if layout is None:
            if index == 1:
                layout = constants.ppLayoutTitleOnly
            else:
                slide = self(index - 1)
                try:
                    layout = slide.CustomLayout  # type: ignore
                except AttributeError:
                    layout = constants.ppLayoutTitleOnly

        if isinstance(layout, int):
            slide = self.api.Add(index, layout)
        else:
            slide = self.api.AddSlide(index, layout)

        return Slide(slide, self)

    @property
    def active(self):
        index = self.app.ActiveWindow.Selection.SlideRange.SlideIndex
        return self(index)


# @dataclass(repr=False)
# class PowerPoint(Base):
#     def __post_init__(self):
#         self.obj = win32com.client.Dispatch("PowerPoint.Application")
#         ensure_modules()

#     @property
#     def presentations(self):
#         return Presentations(self)

#     @property
#     def presentation(self):
#         return self.presentations.active

#     @property
#     def slides(self):
#         return self.presentation.slides

#     @property
#     def slide(self):
#         return self.slides.active

#     @property
#     def shapes(self):
#         return self.slide.shapes

#     @property
#     def tables(self):
#         return self.slide.tables

#     def add_picture(self, *args, **kwargs):
#         return self.slide.shapes.add_picture(*args, **kwargs)

#     def add_frame(self, *args, **kwargs):
#         return self.slide.shapes.add_frame(*args, **kwargs)

#     def add_range(self, *args, **kwargs):
#         return self.slide.shapes.add_range(*args, **kwargs)

#     def add_chart(self, *args, **kwargs):
#         return self.slide.shapes.add_chart(*args, **kwargs)

#     def add_label(self, *args, **kwargs):
#         return self.slide.shapes.add_label(*args, **kwargs)

#     def add_shape(self, *args, **kwargs):
#         return self.slide.shapes.add_shape(*args, **kwargs)

#     def add_table(self, *args, **kwargs):
#         return self.slide.shapes.add_table(*args, **kwargs)
# """
# CustomLayoutに関連するモジュール
# """


# def copy_layout(slide, name=None, replace=True):
#     """指定するスライドのCustomLayoutをコピーして返す．

#     Parameters
#     ----------
#     slide : xlviews.powerpoint.main.Slide
#         スライドオブジェクト
#     name : str, optional
#         CustomLayoutの名前
#     replace : bool, optional
#         スライドのCustomLayoutをコピーしたものに
#         置き換えるか

#     Returns
#     -------
#     layout
#     """
#     layouts = slide.parent.api.SlideMaster.CustomLayouts
#     slide.api.CustomLayout.Copy()
#     layout = layouts.Paste()
#     if name:
#         layout.Name = name
#     if replace:
#         slide.api.CustomLayout = layout
#     return layout
