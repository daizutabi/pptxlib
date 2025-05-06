from pptxlib.core.app import App
from pptxlib.core.presentation import Presentation


def test_add(app: App):
    pr = app.presentations.add()
    assert isinstance(pr, Presentation)
    assert app.presentations.active.name == pr.name


def test_slides(app: App):
    pr = app.presentations.add()
    assert len(pr.slides) == 0


def test_width(app: App):
    pr = app.presentations.add()
    assert pr.width == 960
    assert pr.height == 540
