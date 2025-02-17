from win32com.client import DispatchBaseClass

from pptxlib.core import Slide
from pptxlib.core.base import Collection, Element
from pptxlib.core.shape import Shape, Shapes


def test_shapes(shapes: Shapes):
    assert isinstance(shapes, Collection)
    assert isinstance(shapes, Shapes)
    assert isinstance(shapes.api, DispatchBaseClass)
    assert shapes.name == "Shapes"
    assert len(shapes) == 1  # title shape


def test_shape(shape: Shape):
    assert isinstance(shape, Element)
    assert isinstance(shape, Shape)
    assert isinstance(shape.api, DispatchBaseClass)
    assert shape.name == "Title 1"


def test_text_range(shape: Shape):
    assert isinstance(shape.text_range, DispatchBaseClass)


def test_text(shape: Shape):
    assert shape.text == ""
    shape.text = "Title"
    assert shape.text == "Title"


def test_slide(shape: Shape, slide: Slide):
    assert shape.slide is slide


def test_left(shape: Shape):
    shape.left = 50
    assert shape.left == 50


def test_left_center(shape: Shape):
    shape.left = "center"
    assert round(shape.left + shape.width / 2) == round(shape.slide.width / 2)  # type: ignore


def test_left_neg(shape: Shape):
    shape.left = -50
    assert round(shape.left + shape.width) == round(shape.slide.width - 50)


def test_top(shape: Shape):
    shape.top = 50
    assert shape.top == 50


def test_top_center(shape: Shape):
    shape.top = "center"
    assert round(shape.top + shape.height / 2) == round(shape.slide.height / 2)


def test_top_neg(shape: Shape):
    shape.top = -100
    assert round(shape.top + shape.height) == round(shape.slide.height - 100)


def test_width(shape: Shape):
    assert shape.width > 0
    shape.width = 250
    assert shape.width == 250


def test_height(shape: Shape):
    assert shape.height > 0
    shape.height = 250
    assert shape.height == 250


def test_font_name(shape: Shape):
    shape.font_name = "Meiryo"
    assert shape.font_name == "Meiryo"


def test_font_size(shape: Shape):
    shape.font_size = 32
    assert shape.font_size == 32


def test_font_bold(shape: Shape):
    shape.bold = True
    assert shape.bold is True
    shape.bold = False
    assert shape.bold is False


def test_font_italic(shape: Shape):
    shape.italic = True
    assert shape.italic is True
    shape.italic = False
    assert shape.italic is False


def test_font_color(shape: Shape):
    shape.color = (255, 0, 0)
    assert shape.color == 255
    shape.color = "green"
    assert shape.color == 32768


def test_fill_color(shape: Shape):
    shape.fill_color = (0, 255, 0)
    assert shape.fill_color == 255 * 256


def test_line_color(shape: Shape):
    shape.line_color = (0, 0, 255)
    assert shape.line_color == 16777215


def test_line_width(shape: Shape):
    shape.line_weight = 2
    assert shape.line_weight == 2


def test_set(shape: Shape):
    shape.set(
        font="Times",
        size=10,
        bold=True,
        italic=True,
        color="blue",
        fill_color="red",
        line_weight=3,
        line_color="green",
    )
    assert shape.font_name == "Times"
    assert shape.font_size == 10
    assert shape.bold is True
    assert shape.italic is True
    assert shape.color == 16711680
    assert shape.fill_color == 255
    assert shape.line_weight == 3
    assert shape.line_color == 32768


def test_add(shapes: Shapes):
    shape = shapes.add("Oval", 100, 100, 40, 60)
    assert shape.text == ""
    assert shape.left == 100
    assert shape.top == 100
    assert shape.width == 40
    assert shape.height == 60
    assert shape.api.Parent.__class__.__name__ == "_Slide"
    assert shape.parent.__class__.__name__ == "Slide"
    shape.delete()


def test_add_label(shapes: Shapes):
    shape = shapes.add_label("ABC", 100, 100)
    assert shape.text == "ABC"
    assert shape.left == 100
    assert shape.top == 100
    width = shape.width
    height = shape.height
    shape.text = "ABC ABC"
    assert width < shape.width
    assert height == shape.height
    assert shape.api.Parent.__class__.__name__ == "_Slide"
    assert shape.parent.__class__.__name__ == "Slide"
    shape.delete()


def test_add_label_auto_size_false(shapes: Shapes):
    shape = shapes.add_label("ABC", 100, 100, 200, 300, auto_size=False)
    assert shape.width == 200
    assert shape.height == 300
    shape.delete()


def test_add_table(shapes: Shapes):
    shape = shapes.add_table(2, 3, 100, 100, 240, 360)
    assert isinstance(shape, Shape)
    assert shape.api.Parent.__class__.__name__ == "_Slide"
    assert shape.parent.__class__.__name__ == "Slide"
    assert shape.api.Table.__class__.__name__ == "Table"
    assert shape.api.Table.Parent.__class__.__name__ == "Shape"
    shape.delete()


def test_slides_repr(shapes: Shapes):
    assert repr(shapes) == "<Shapes>"


def test_slide_repr(shape: Shape):
    assert repr(shape) == "<Shape [Title 1]>"


def test_slide_oval_repr(shapes: Shapes):
    shape = shapes.add("Oval", 100, 100, 40, 60)
    assert repr(shape) == "<Shape [Oval 2]>"
    shape.delete()


def test_slides_parent(shapes: Shapes):
    assert shapes.api.Parent.__class__.__name__ == "_Slide"
    assert shapes.parent.__class__.__name__ == "Slide"


def test_slide_parent(shape: Shape, shapes: Shapes):
    assert shape.api.Parent.__class__.__name__ == "_Slide"
    assert shape.parent.__class__.__name__ == "Slide"
    assert shapes(1).parent.__class__.__name__ == "Slide"
