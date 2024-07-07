from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, ClassVar, Generic, TypeVar

if TYPE_CHECKING:
    from collections.abc import Iterator

    from win32com.client import CoClassBaseClass, DispatchBaseClass


@dataclass(repr=False)
class Base:
    api: DispatchBaseClass = field(init=False)
    app: DispatchBaseClass = field(init=False)

    def __repr__(self):
        clsname = self.__class__.__name__
        return f"<{clsname}>"

    @property
    def name(self):
        try:
            return self.api.Name

        except AttributeError:
            return self.__class__.__name__


@dataclass(repr=False)
class Element(Base):
    api: CoClassBaseClass
    parent: Base

    def __post_init__(self):
        self.app = self.parent.app

    def __repr__(self):
        clsname = self.__class__.__name__
        return f"<{clsname} [{self.name}]>"

    def select(self):
        self.api.Select()

    def delete(self):
        self.api.Delete()


SomeElement = TypeVar("SomeElement", bound=Element)


@dataclass(repr=False)
class Collection(Base, Generic[SomeElement]):
    parent: Base
    type: ClassVar[type[Element]] = field(init=False)

    def __post_init__(self):
        self.api = getattr(self.parent.api, self.__class__.__name__)
        self.app = self.parent.app

    def __len__(self) -> int:
        return self.api.Count

    def __call__(self, index: int | None = None) -> SomeElement:
        if index is None:
            index = len(self) + 1

        return self.type(self.api(index), self)  # type: ignore

    def __iter__(self) -> Iterator[SomeElement]:
        for index in range(len(self)):
            yield self(index + 1)

    def __getitem__(self, index) -> SomeElement | list[SomeElement]:
        if isinstance(index, slice):
            return list(self)[index]

        if index < 0:
            index = len(self) + index

        return self(index + 1)
