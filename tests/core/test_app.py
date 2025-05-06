from pptxlib.core.app import App
from pptxlib.core.presentation import Presentations


def test_presentations(app: App):
    presentations = app.presentations
    assert isinstance(presentations, Presentations)
