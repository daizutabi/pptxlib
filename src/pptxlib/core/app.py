from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Self

import win32com.client

from pptxlib.client import ensure_modules

from .base import Base
from .presentation import Presentations

if TYPE_CHECKING:
    from win32com.client import DispatchBaseClass


@dataclass(repr=False)
class App(Base):
    api: DispatchBaseClass = field(init=False)

    def __post_init__(self) -> None:
        ensure_modules()
        self.api = win32com.client.Dispatch("PowerPoint.Application")  # type: ignore
        self.app = self.api

    @property
    def presentations(self) -> Presentations:
        return Presentations(self.api.Presentations, self)

    def quit(self) -> None:
        self.api.Quit()

    def __enter__(self) -> Self:
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:  # noqa: ANN001
        self.quit()
