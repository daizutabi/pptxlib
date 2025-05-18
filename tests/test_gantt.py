from datetime import datetime

import pytest


def test_date_index_month():
    from pptxlib.gantt import date_index

    index = date_index("month", datetime(2024, 4, 1), datetime(2025, 9, 10))
    assert len(index) == 18
    assert index[0] == datetime(2024, 4, 1)
    assert index[-1] == datetime(2025, 9, 1)


def test_date_index_week():
    from pptxlib.gantt import date_index

    index = date_index("week", datetime(2025, 5, 7), datetime(2025, 6, 10))
    assert len(index) == 6
    assert index[0] == datetime(2025, 5, 5)
    assert index[-1] == datetime(2025, 6, 9)


def test_date_index_day():
    from pptxlib.gantt import date_index

    index = date_index("day", datetime(2025, 5, 7), datetime(2025, 5, 14))
    assert len(index) == 8
    assert index[0] == datetime(2025, 5, 7)
    assert index[-1] == datetime(2025, 5, 14)


def test_date_index_error():
    from pptxlib.gantt import date_index

    with pytest.raises(ValueError):
        date_index("invalid", datetime(2025, 5, 7), datetime(2025, 5, 14))


def test_name_month():
    from pptxlib.gantt import GanttFrame

    gantt = GanttFrame("month", datetime(2024, 4, 1), datetime(2025, 9, 10))
    assert gantt.name == "2024/04/01-2025/09/01-month"


def test_name_week():
    from pptxlib.gantt import GanttFrame

    gantt = GanttFrame("week", datetime(2025, 4, 1), datetime(2025, 9, 10))
    assert gantt.name == "2025/03/31-2025/09/08-week"


def test_name_day():
    from pptxlib.gantt import GanttFrame

    gantt = GanttFrame("day", datetime(2025, 4, 1), datetime(2025, 9, 10))
    assert gantt.name == "2025/04/01-2025/09/10-day"
