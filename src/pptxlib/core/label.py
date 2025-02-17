from __future__ import annotations

from dataclasses import dataclass
from typing import Self

from pptxlib.core.shape import Shape


@dataclass(repr=False)
class Label(Shape):
    def set(
        self,
        name: str | None = None,
        size: float | None = None,
        bold: bool | None = None,
        italic: bool | None = None,
        color: int | str | tuple[int, int, int] | None = None,
    ) -> Self:
        self.font.set(name, size, bold, italic, color)
        return self
