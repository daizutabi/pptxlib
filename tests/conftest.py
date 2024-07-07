import pytest

from pptxlib.app import PowerPoint, Presentation, Presentations, Slide, Slides
from pptxlib.shapes import Shape, Shapes


@pytest.fixture(scope="session")
def pp():
    pp = PowerPoint()
    yield pp
    pp.quit()


@pytest.fixture(scope="session")
def prs(pp: PowerPoint):
    return pp.presentations


@pytest.fixture
def pr(prs: Presentations):
    pr = prs.add()
    yield pr
    pr.close()


@pytest.fixture
def pr_list(prs: Presentations):
    pr = [prs.add() for _ in range(3)]
    yield pr
    for p in pr:
        p.close()


@pytest.fixture
def slides(pr: Presentation):
    return pr.slides


@pytest.fixture
def slide(slides: Slides):
    slide = slides.add()
    yield slide
    slide.delete()


@pytest.fixture
def shapes(slide: Slide):
    return slide.shapes


@pytest.fixture
def title_shape(shapes: Shapes):
    return shapes(1)
