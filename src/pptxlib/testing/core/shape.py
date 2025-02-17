from __future__ import annotations

from typing import TYPE_CHECKING

from win32com.client import constants

from pptxlib.core.font import Font
from pptxlib.testing.common import create_slide

if TYPE_CHECKING:
    from pptxlib.core.presentation import Presentation


def main():
    slide = create_slide()
    table = slide.shapes.add_table(2, 3, 100, 100, 100, 100)
    print(table[0].cells[0].borders)
    b = table[0].cells.borders[0]
    print(b)
    print(b.api)
    print(b.parent)
    print(b.collection)
    # print(table[0].cells.borders["bottom"])


if __name__ == "__main__":
    main()
