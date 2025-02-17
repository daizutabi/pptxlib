from __future__ import annotations

from functools import cache
from typing import TYPE_CHECKING

from pywintypes import com_error

from pptxlib.core.app import App

if TYPE_CHECKING:
    from pptxlib.core.presentation import Presentation
    from pptxlib.core.slide import Slide


@cache
def is_app_available() -> bool:
    try:
        with App():
            pass
    except com_error:
        return False

    return True


def create_presentation() -> Presentation:
    app = App()

    for pr in app.presentations:
        pr.close()

    return app.presentations.add()


def create_slide() -> Slide:
    pr = create_presentation()
    return pr.slides.add()
