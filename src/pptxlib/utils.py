import matplotlib.colors
from win32com.client import constants


def to_bool(x: int) -> bool:
    return x == constants.msoTrue


def rgb(color: int | str | tuple[int, int, int]):
    """Color function.

    Args:
        color: Color.

    Examples:
        >>> rgb(4)
        4
        >>> rgb((100, 200, 40))
        2672740
        >>> rgb("pink")
        13353215
        >>> rgb("#123456")
        5649426
    """
    if isinstance(color, int):
        return color

    if isinstance(color, str):
        color = str(matplotlib.colors.cnames.get(color, color))

        if not color.startswith("#") or len(color) != 7:  # noqa: PLR2004
            raise ValueError

        red = int(color[1:3], 16)
        green = int(color[3:5], 16)
        blue = int(color[5:7], 16)

    else:
        red, green, blue = color

    return red + green * 256 + blue * 256 * 256
