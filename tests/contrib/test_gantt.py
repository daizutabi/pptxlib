from datetime import datetime

import pytest

from pptxlib.app import is_powerpoint_available
from pptxlib.contrib.gantt import GanttChart
from pptxlib.presentation import Presentations

pytestmark = pytest.mark.skipif(
    not is_powerpoint_available(),
    reason="PowerPoint is not available",
)


@pytest.fixture(scope="module")
def gantt(prs: Presentations):
    pr = prs.add(with_window=False)
    gantt = GanttChart("2025/5/21", "2025/6/10", kind="week")
    slide = pr.slides.add()
    gantt.add_table(slide, 20, 150)
    yield gantt
    pr.close()


def test_table(gantt: GanttChart):
    assert gantt.table.left == 20
    assert gantt.table.top == 150
    assert gantt.table.width == 920
    assert gantt.table.height == 240


def test_add(gantt: GanttChart):
    s1 = gantt.add(datetime(2025, 5, 21), 20, color="red")
    assert 183 < s1.left < 184
    assert 207 < s1.top < 209


def test_add_str(gantt: GanttChart):
    s1 = gantt.add("2025/05/21", 40, color="red")
    assert 183 < s1.left < 184
    assert 227 < s1.top < 229
