import pytest
from win32com.client import DispatchBaseClass

from pptxlib.core.app import is_app_available
from pptxlib.core.base import Collection, Element
from pptxlib.core.shape import Shape, Shapes
from pptxlib.core.slide import Slide
from pptxlib.core.table import Table

pytestmark = pytest.mark.skipif(
    not is_app_available(),
    reason="PowerPoint is not available",
)


def test_shapes(shapes: Shapes):
    assert isinstance(shapes, Collection)
    assert isinstance(shapes, Shapes)
    assert isinstance(shapes.api, DispatchBaseClass)
    assert len(shapes) == 1  # title shape


def test_shapes_repr(shapes: Shapes):
    assert repr(shapes) == "<Shapes (1)>"


@pytest.fixture
def shape(shapes: Shapes):
    return shapes[0]


def test_title(shapes: Shapes, shape: Shape):
    assert shapes.title.name == shape.name


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


def test_parent(shape: Shape, slide: Slide):
    assert shape.parent is slide


def test_left(shape: Shape):
    shape.left = 50
    assert shape.left == 50


def test_left_center(shape: Shape, slide: Slide):
    shape.left = "center"
    assert round(shape.left + shape.width / 2) == round(slide.width / 2)  # type: ignore


def test_left_neg(shape: Shape, slide: Slide):
    shape.left = -50
    assert round(shape.left + shape.width) == round(slide.width - 50)


def test_top(shape: Shape):
    shape.top = 50
    assert shape.top == 50


def test_top_center(shape: Shape, slide: Slide):
    shape.top = "center"
    assert round(shape.top + shape.height / 2) == round(slide.height / 2)


def test_top_neg(shape: Shape, slide: Slide):
    shape.top = -100
    assert round(shape.top + shape.height) == round(slide.height - 100)


def test_width(shape: Shape):
    assert shape.width > 0
    shape.width = 250
    assert shape.width == 250


def test_height(shape: Shape):
    assert shape.height > 0
    shape.height = 250
    assert shape.height == 250


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
    table = shapes.add_table(2, 3, 100, 100, 240, 360)
    assert isinstance(table, Table)
    table.delete()


def test_shape_repr(shape: Shape):
    assert repr(shape) == "<Shape [Title 1]>"


def test_shape_oval_repr(shapes: Shapes):
    shape = shapes.add("Oval", 100, 100, 40, 60)
    assert repr(shape) == "<Shape [Oval 2]>"
    shape.delete()


def test_shapes_parent(shapes: Shapes):
    assert shapes.api.Parent.__class__.__name__ == "_Slide"
    assert shapes.parent.__class__.__name__ == "Slide"


def test_shape_parent(shape: Shape, shapes: Shapes):
    assert shape.api.Parent.__class__.__name__ == "_Slide"
    assert shape.parent.__class__.__name__ == "Slide"
    assert shapes[0].parent.__class__.__name__ == "Slide"
