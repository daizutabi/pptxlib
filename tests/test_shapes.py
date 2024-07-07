from win32com.client import DispatchBaseClass

from pptxlib.app import Slide
from pptxlib.core import Collection, Element
from pptxlib.shapes import Shape, Shapes


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
    assert shape.left > 0
    shape.left = 50
    assert shape.left == 50


def test_top(shape: Shape):
    assert shape.top > 0
    shape.top = 50
    assert shape.top == 50


def test_width(shape: Shape):
    assert shape.width > 0
    shape.width = 250
    assert shape.width == 250


def test_height(shape: Shape):
    assert shape.height > 0
    shape.height = 250
    assert shape.height == 250


def test_repr_slides(shapes: Shapes):
    assert repr(shapes) == "<Shapes>"


def test_repr_slide(shape: Shape):
    assert repr(shape) == "<Shape [Title 1]>"
