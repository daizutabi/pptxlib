from datetime import datetime

from pptxlib.contrib.gantt import GanttFrame, date_index, fiscal_year


def test_date_index_month_unit():
    index = date_index("month", datetime(2024, 4, 10), datetime(2025, 9, 10))
    assert len(index) == 18
    assert index[0] == datetime(2024, 4, 1)
    assert index[-1] == datetime(2025, 9, 1)


def test_date_index_week_unit():
    index = date_index("week", datetime(2025, 5, 7), datetime(2025, 6, 10))
    assert len(index) == 6
    assert index[0] == datetime(2025, 5, 5)
    assert index[-1] == datetime(2025, 6, 9)


def test_date_index_day_unit():
    index = date_index("day", datetime(2025, 5, 7), datetime(2025, 5, 14))
    assert len(index) == 8
    assert index[0] == datetime(2025, 5, 7)
    assert index[-1] == datetime(2025, 5, 14)


def test_fiscal_year():
    assert fiscal_year(datetime(2025, 1, 1)) == "FY2024"
    assert fiscal_year(datetime(2025, 4, 1)) == "FY2025"


def test_ganttframe_name():
    gf = GanttFrame("week", datetime(2025, 4, 1), datetime(2025, 9, 10))
    assert gf.name == "2025/03/31-2025/09/08-week"
