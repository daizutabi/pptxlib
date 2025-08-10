from datetime import datetime

import pytest

from pptxlib.contrib.gantt import GanttFrame, date_index, fiscal_year


def test_date_index_month():
    from pptxlib.contrib.gantt import date_index

    index = date_index(datetime(2024, 4, 1), datetime(2025, 9, 10), kind="month")
    assert len(index) == 18
    assert index[0] == datetime(2024, 4, 1)
    assert index[-1] == datetime(2025, 9, 1)


def test_date_index_week():
    from pptxlib.contrib.gantt import date_index

    index = date_index(datetime(2025, 5, 7), datetime(2025, 6, 10), kind="week")
    assert len(index) == 6
    assert index[0] == datetime(2025, 5, 5)
    assert index[-1] == datetime(2025, 6, 9)


def test_date_index_day():
    from pptxlib.contrib.gantt import date_index

    index = date_index(datetime(2025, 5, 7), datetime(2025, 5, 14), kind="day")
    assert len(index) == 8
    assert index[0] == datetime(2025, 5, 7)
    assert index[-1] == datetime(2025, 5, 14)


def test_date_index_error():
    from pptxlib.contrib.gantt import date_index

    with pytest.raises(ValueError):
        date_index(datetime(2025, 5, 7), datetime(2025, 5, 14), kind="invalid")


def test_date_index_month_unit():
    index = date_index(datetime(2024, 4, 10), datetime(2025, 9, 10), kind="month")
    assert len(index) == 18
    assert index[0] == datetime(2024, 4, 1)
    assert index[-1] == datetime(2025, 9, 1)


def test_date_index_week_unit():
    index = date_index(datetime(2025, 5, 7), datetime(2025, 6, 10), kind="week")
    assert len(index) == 6
    assert index[0] == datetime(2025, 5, 5)
    assert index[-1] == datetime(2025, 6, 9)


def test_date_index_day_unit():
    index = date_index(datetime(2025, 5, 7), datetime(2025, 5, 14), kind="day")
    assert len(index) == 8
    assert index[0] == datetime(2025, 5, 7)
    assert index[-1] == datetime(2025, 5, 14)


def test_fiscal_year():
    assert fiscal_year(datetime(2025, 1, 1)) == "FY2024"
    assert fiscal_year(datetime(2025, 4, 1)) == "FY2025"


def test_ganttframe_name():
    gf = GanttFrame(datetime(2025, 4, 1), datetime(2025, 9, 10), kind="week")
    assert gf.name == "2025/03/31-2025/09/08-week"


def test_name_month():
    from pptxlib.contrib.gantt import GanttFrame

    gantt = GanttFrame(datetime(2024, 4, 1), datetime(2025, 9, 10), kind="month")
    assert gantt.name == "2024/04/01-2025/09/01-month"


def test_name_week():
    from pptxlib.contrib.gantt import GanttFrame

    gantt = GanttFrame(datetime(2025, 4, 1), datetime(2025, 9, 10), kind="week")
    assert gantt.name == "2025/03/31-2025/09/08-week"


def test_name_day():
    from pptxlib.contrib.gantt import GanttFrame

    gantt = GanttFrame(datetime(2025, 4, 1), datetime(2025, 9, 10), kind="day")
    assert gantt.name == "2025/04/01-2025/09/10-day"


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
