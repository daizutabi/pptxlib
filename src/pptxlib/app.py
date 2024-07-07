from __future__ import annotations

from dataclasses import dataclass
from typing import ClassVar

import win32com.client
from win32com.client import constants

from pptxlib.client import ensure_modules
from pptxlib.core import Base, Collection, Element
from pptxlib.shapes import Shapes


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
    def close(self):
        self.api.Close()

    @property
    def slides(self):
        return Slides(self)


@dataclass(repr=False)
class Presentations(Collection[Presentation]):
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
    @property
    def shapes(self) -> Shapes:
        return Shapes(self)

    @property
    def title(self) -> str:
        return self.shapes(1).text if len(self.shapes) else ""

    @title.setter
    def title(self, text):
        if len(self.shapes):
            self.shapes(1).text = text

    # @property
    # def tables(self):
    #     return Tables(self)


class Slides(Collection[Slide]):
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
