from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

from pptxlib.core.base import Element

if TYPE_CHECKING:
    from .table import Borders


@dataclass(repr=False)
class LineFormat(Element):
    parent: Borders
    collection: Borders


# def get_borders(
#     cell: Cell | CellRange,
#     border_type: Literal["bottom","left","right","top"],
#     width: float = 1,
#     color: int | str | tuple[int, int, int] = 0,
#     line_style: Literal["-", "--"] = "-",
#     *,
#     visible: bool = True,
# ):
#     border_type_int = getattr(constants, "ppBorder" + border_type[0].upper() + border_type[1:])
#     border = cell.api.Borders(border_type_int)
#     border.Visible = visible

#     if not visible:
#         return

#     border.Weight = width
#     border.ForeColor.RGB = color

#     if line_style == "--":
#         border.DashStyle = constants.msoLineDash"
