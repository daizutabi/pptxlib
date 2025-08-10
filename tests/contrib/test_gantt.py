from datetime import datetime

import pytest

from pptxlib.app import is_powerpoint_available
from pptxlib.contrib.gantt import GanttChart
from pptxlib.presentation import Presentations

pytestmark = pytest.mark.skipif(
    not is_powerpoint_available(),
    reason="PowerPoint is not available",
)


def test_date_index_month():
    from pptxlib.contrib.gantt import date_index

    index = date_index("month", datetime(2024, 4, 1), datetime(2025, 9, 10))
    assert len(index) == 18
    assert index[0] == datetime(2024, 4, 1)
    assert index[-1] == datetime(2025, 9, 1)


def test_date_index_week():
    from pptxlib.contrib.gantt import date_index

    index = date_index("week", datetime(2025, 5, 7), datetime(2025, 6, 10))
    assert len(index) == 6
    assert index[0] == datetime(2025, 5, 5)
    assert index[-1] == datetime(2025, 6, 9)


def test_date_index_day():
    from pptxlib.contrib.gantt import date_index

    index = date_index("day", datetime(2025, 5, 7), datetime(2025, 5, 14))
    assert len(index) == 8
    assert index[0] == datetime(2025, 5, 7)
    assert index[-1] == datetime(2025, 5, 14)


def test_date_index_error():
    from pptxlib.contrib.gantt import date_index

    with pytest.raises(ValueError):
        date_index("invalid", datetime(2025, 5, 7), datetime(2025, 5, 14))


def test_name_month():
    from pptxlib.contrib.gantt import GanttFrame

    gantt = GanttFrame("month", datetime(2024, 4, 1), datetime(2025, 9, 10))
    assert gantt.name == "2024/04/01-2025/09/01-month"


def test_name_week():
    from pptxlib.contrib.gantt import GanttFrame

    gantt = GanttFrame("week", datetime(2025, 4, 1), datetime(2025, 9, 10))
    assert gantt.name == "2025/03/31-2025/09/08-week"


def test_name_day():
    from pptxlib.contrib.gantt import GanttFrame

    gantt = GanttFrame("day", datetime(2025, 4, 1), datetime(2025, 9, 10))
    assert gantt.name == "2025/04/01-2025/09/10-day"


@pytest.fixture(scope="module")
def gantt(prs: Presentations):
    pr = prs.add(with_window=False)
    gantt = GanttChart("week", "2025/5/21", "2025/6/10")
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


@pytest.mark.parametrize("date", ["2025/05/21", "2025-5-21", "2025.5.21"])
def test_strptime(date: str):
    from pptxlib.contrib.gantt import strptime

    d = strptime(date)
    assert d.year == 2025
    assert d.month == 5
    assert d.day == 21


def test_strptime_invalid_separator():
    from pptxlib.contrib.gantt import strptime

    with pytest.raises(ValueError):
        strptime("2025!05!21")


def test_strptime_invalid_count():
    from pptxlib.contrib.gantt import strptime

    with pytest.raises(ValueError):
        strptime("2025/05/21/x")
