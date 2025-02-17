from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, ClassVar

from win32com.client import constants

from pptxlib.core.base import Collection, Element
from pptxlib.core.shape import Shapes

if TYPE_CHECKING:
    from .presentation import Presentation


@dataclass(repr=False)
class Slide(Element):
    parent: Presentation
    collection: Slides

    @property
    def shapes(self) -> Shapes:
        return Shapes(self.api.Shapes, self)

    # @property
    # def tables(self) -> Tables:
    #     return Tables(self.api.Shapes, self)

    @property
    def title(self) -> str:
        return self.shapes.title.text if len(self.shapes) else ""

    @title.setter
    def title(self, text: str) -> None:
        if len(self.shapes):
            self.shapes.title.text = text

    @property
    def width(self) -> float:
        return self.parent.width

    @property
    def height(self) -> float:
        return self.parent.height


@dataclass(repr=False)
class Slides(Collection[Slide]):
    parent: Presentation
    type: ClassVar[type[Element]] = Slide

    def add(self, index: int | None = None, layout: int | str | None = None) -> Slide:
        if index is None:
            index = len(self)

        if isinstance(layout, str):
            layout = getattr(constants, f"ppLayout{layout}")
        elif layout is None:
            if index == 0:
                layout = constants.ppLayoutTitleOnly
            else:
                slide = self[index - 1]
                try:
                    layout = slide.api.CustomLayout
                except AttributeError:
                    layout = constants.ppLayoutTitleOnly

        if isinstance(layout, int):
            slide = self.api.Add(index + 1, layout)
        else:
            slide = self.api.AddSlide(index + 1, layout)

        return Slide(slide, self.parent, self)

    @property
    def active(self) -> Slide:
        index = self.app.ActiveWindow.Selection.SlideRange.SlideIndex - 1
        return self[index]
