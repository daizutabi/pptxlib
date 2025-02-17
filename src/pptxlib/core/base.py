from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, ClassVar, Generic, TypeVar

if TYPE_CHECKING:
    from collections.abc import Iterator

    from win32com.client import CoClassBaseClass, DispatchBaseClass


@dataclass(repr=False)
class Base:
    api: DispatchBaseClass | CoClassBaseClass
    app: DispatchBaseClass = field(init=False)

    def __repr__(self) -> str:
        clsname = self.__class__.__name__
        return f"<{clsname}>"

    @property
    def name(self) -> str:
        return self.api.Name

    @name.setter
    def name(self, value: str) -> None:
        self.api.Name = value


@dataclass(repr=False)
class Element(Base):
    parent: Element
    collection: Collection

    def __post_init__(self) -> None:
        self.app = self.parent.app

    def __repr__(self) -> str:
        clsname = self.__class__.__name__
        return f"<{clsname} [{self.name}]>"

    def select(self) -> None:
        self.api.Select()

    def delete(self) -> None:
        self.api.Delete()


E = TypeVar("E", bound=Element)


@dataclass(repr=False)
class Collection(Base, Generic[E]):
    parent: Element
    type: ClassVar[type[Element]]

    def __post_init__(self) -> None:
        self.app = self.parent.app

    def __len__(self) -> int:
        return self.api.Count

    def __repr__(self) -> str:
        clsname = self.__class__.__name__
        return f"<{clsname} ({len(self)})>"

    def __getitem__(self, index: int) -> E:
        if index < 0:
            index = len(self) + index

        return self.type(self.api(index + 1), self.parent, self)  # type: ignore

    def __iter__(self) -> Iterator[E]:
        yield from [self[index] for index in range(len(self))]  # list due to deletion
