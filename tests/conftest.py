import pytest

from pptxlib.app import PowerPoint, Presentation, Presentations, Slide, Slides
from pptxlib.shapes import Shapes
from pptxlib.tables import Table, Tables


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
def shape(shapes: Shapes):
    return shapes.title


@pytest.fixture
def tables(slide: Slide):
    return slide.tables


@pytest.fixture
def table(tables: Tables):
    table = tables.add(2, 3, 100, 100, 200, 200)
    yield table
    table.delete()


@pytest.fixture
def rows(table: Table):
    return table.rows


@pytest.fixture
def columns(table: Table):
    return table.columns
