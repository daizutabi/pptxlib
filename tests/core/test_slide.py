import pytest

from pptxlib.core.slide import Slide, Slides


@pytest.fixture
def slide(slides: Slides):
    return slides.add()


def test_active(slides: Slides, slide: Slide):
    assert slides.active.name == slide.name


def test_width(slide: Slide):
    assert slide.width == 960


def test_height(slide: Slide):
    assert slide.height == 540


def test_title(slide: Slide):
    slide.title = "Title"
    assert slide.title == "Title"
