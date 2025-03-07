import pytest

from pptxlib.core.app import App
from pptxlib.core.presentation import Presentation, Presentations
from pptxlib.core.shape import Shapes
from pptxlib.core.slide import Slide, Slides
from pptxlib.core.table import Rows, Table


@pytest.fixture(scope="session")
def app():
    app = App()
    yield app
    app.quit()


@pytest.fixture(scope="session")
def prs(app: App):
    return app.presentations


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
def rows(table: Table):
    return table.rows


@pytest.fixture
def columns(table: Table):
    return table.columns


@pytest.fixture
def cell(table: Table):
    return table.cell(1, 1)


@pytest.fixture
def cell_range(rows: Rows):
    return rows(1).cells
