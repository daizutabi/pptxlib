from __future__ import annotations

from datetime import datetime
from enum import Enum

# from itertools import product
# import pandas as pd
from dateutil.relativedelta import relativedelta

from pptxlib.core.presentation import Presentation
from pptxlib.core.slide import Layout
from pptxlib.core.table import Table

# from win32com.client import constants

# import xlviews as xv
# from xlviews.powerpoint.connector import add_connector
# from xlviews.powerpoint.layout import copy_layout
# from xlviews.powerpoint.main import Shape
# from xlviews.powerpoint.style import set_fill
# from xlviews.powerpoint.table import create_table
# from xlviews.utils import rgb


def date_index(kind: str, start: datetime, end: datetime) -> list[datetime]:
    if kind in ["month", "monthly"]:
        start = datetime(start.year, start.month, 1)
        end = datetime(end.year, end.month, 1)
        delta = relativedelta(end, start)
        n = 12 * delta.years + delta.months
        return [start + relativedelta(months=k) for k in range(n + 1)]

    if kind in ["week", "weekly"]:
        start -= relativedelta(days=start.weekday())
        end -= relativedelta(days=end.weekday())
        n = (end - start).days // 7
        return [start + relativedelta(days=7 * k) for k in range(n + 1)]

    if kind in ["day", "daily"]:
        n = (end - start).days
        return [start + relativedelta(days=k) for k in range(n + 1)]

    msg = f"Unsupported kind: {kind}"
    raise ValueError(msg)


def fiscal_year(date: datetime) -> str:
    if 1 <= date.month <= 3:
        return f"FY{date.year - 1}"

    return f"FY{date.year}"


class GanttKind(Enum):
    MONTH = "month"
    WEEK = "week"
    DAY = "day"


class GanttFrame:
    kind: GanttKind
    date_index: list[datetime]
    columns: list[list[str]]

    def __init__(self, kind: str, start: datetime, end: datetime) -> None:
        self.date_index = date_index(kind, start, end)

        years = [fiscal_year(date) for date in self.date_index]
        months = [str(date.month) for date in self.date_index]
        days = [str(date.day) for date in self.date_index]

        if kind in ["month", "monthly"]:
            self.columns = [years, months]
            self.kind = GanttKind.MONTH
        elif kind in ["week", "weekly"]:
            self.columns = [years, months, days]
            self.kind = GanttKind.WEEK
        elif kind in ["day", "daily"]:
            self.columns = [years, months, days]
            self.kind = GanttKind.DAY
        else:
            raise NotImplementedError

    @property
    def name(self) -> str:
        start = self.date_index[0].strftime("%Y/%m/%d")
        end = self.date_index[-1].strftime("%Y/%m/%d")
        return f"{start}-{end}-{self.kind.value}"


class GanttChart:
    frame: GanttFrame
    layout: Layout
    table: Table

    def __init__(
        self,
        kind: str,
        start: datetime,
        end: datetime,
        pr: Presentation,
        left: float,
        top: float,
        right: float | None = None,
        bottom: float | None = None,
        index_width: float = 80,
    ) -> None:
        self.frame = GanttFrame(kind, start, end)
        self.layout = pr.layouts.add(self.frame.name)

        if right is None:
            right = left
        if bottom is None:
            bottom = top

        self.table = self.layout.shapes.add_table(
            num_rows=len(self.frame.columns),
            num_columns=len(self.frame.columns[0]) + 1,
            left=left,
            top=top,
            width=self.layout.width - left - right,
            height=self.layout.height - top - bottom,
        )
        self.table.clear()

        # for cell in row.cells:
        #     cell.shape.fill.api.Visible = False


#         columns_name = self.how in ['year', 'yearly']
#         shape = create_table(layout.Shapes, self.frame,
#                              columns_name=columns_name,
#                              left=left_margin, top=top_margin,
#                              width=width, height=height, preclean=False)
#         shape.api.Name = self.name
#         self.table = shape.table

#         table = self.table.api
#         table.FirstRow = False
#         table.HorizBanding = False
#         nrows = len(self.frame.columns.names)
#         ncols = len(self.frame.index.names)
#         for row, column in product(range(nrows), range(ncols)):
#             cell = table.Cell(row + 1, column + 1)
#             cell.Shape.Fill.Visible = False

#         columns_level = len(self.frame.columns.names)
#         if columns_level >= 2:
#             for cell in self.table.row(2):
#                 cell.shape.size = 10
#         if columns_level == 3:
#             for cell in self.table.row(3):
#                 cell.shape.size = 8

#         self.table.columns(1).width = index_width
#         column_width = (width - index_width) / len(self.frame.columns)
#         for k in range(len(self.frame.columns)):
#             self.table.columns(k + 2).width = column_width

#         columns_height = sum(self.table.rows(k + 1).height
#                              for k in range(columns_level))
#         index_height = height - columns_height
#         self.table.rows(len(self.frame) + columns_level).height = index_height

#         for column in range(len(self.frame.columns)):
#             color = (rgb(246, 250, 252) if column % 2 else rgb(246, 252, 246))
#             set_fill(table, (columns_level, column + 2),
#                      (columns_level + 1, column + 2), color)
#             cell = self.table.cell(columns_level, column + 2)
#             text_frame2 = cell.shape.api.TextFrame2
#             text_frame2.MarginBottom = 0
#             text_frame2.MarginTop = 0
#             text_frame2.MarginLeft = 0
#             text_frame2.MarginRight = 0

#         self.calc_scale()

#     def calc_scale(self):
#         columns_level = len(self.frame.columns.names)
#         self.left = self.table.cell(1, 2).left
#         self.top = self.table.cell(columns_level + 1, 2).top
#         end = self.table.left + self.table.parent.width
#         self.day_width = (end - self.left) / ((self.end - self.start).days + 1)
#         self.bottom = self.table.top + self.table.parent.height
#         self.height = self.bottom - self.top

#     def to_position(self, date, offset=0.5):
#         return ((date - self.start).days + offset) * self.day_width + self.left

#     def to_date(self, position, offset=0.5):
#         days = round((position - self.left) / gc.day_width - offset)
#         return self.start + relativedelta(days=days)

#     def add_point(self, date, y, width=10, height=None, text='day', shape='r',
#                   size=10, **kwargs):
#         if isinstance(shape, str):
#             shape = {'r': 5, 'd': 4, '^': 7, 'o': 9}[shape]
#         if height is None:
#             height = width
#         left = self.to_position(date) - width / 2
#         top = y * self.height + self.top - height / 2

#         if text == 'day':
#             text = str(date.day)

#         shape = self.slide.shapes.add_shape(shape, left, top, width, height,
#                                             text=text, italic=False, size=size,
#                                             **kwargs)
#         shape.api.Name = 'gantt_point'
#         shape.api.Fill.Solid()
#         shape.api.Shadow.Visible = False

#         if text:
#             text_frame2 = shape.api.TextFrame2
#             text_frame2.VerticalAnchor = constants.msoAnchorMiddle
#             paragraph_format = text_frame2.TextRange.ParagraphFormat
#             paragraph_format.Alignment = constants.msoAlignCenter
#             text_frame2.MarginBottom = 0
#             text_frame2.MarginTop = 0
#             text_frame2.MarginLeft = 0
#             text_frame2.MarginRight = 0
#             shape.api.TextFrame.TextRange.Font.Shadow = False

#         return shape

#     def add_line(self, start, end, y, height=10, text=None, **kwargs):
#         shape = 5
#         left = self.to_position(start, offset=0)
#         right = self.to_position(end, offset=1)
#         width = right - left
#         top = y * self.height + self.top - height / 2

#         shape = self.slide.shapes.add_shape(shape, left, top, width, height,
#                                             text=text, italic=False, **kwargs)
#         shape.api.Name = 'gantt_line'
#         shape.api.Fill.Solid()
#         shape.api.Shadow.Visible = False

#         if text:
#             text_frame2 = shape.api.TextFrame2
#             text_frame2.VerticalAnchor = constants.msoAnchorMiddle
#             paragraph_format = text_frame2.TextRange.ParagraphFormat
#             paragraph_format.Alignment = constants.msoAlignCenter
#             text_frame2.MarginBottom = 0
#             text_frame2.MarginTop = 0
#             text_frame2.MarginLeft = 0
#             text_frame2.MarginRight = 0
#             shape.api.TextFrame.TextRange.Font.Shadow = False

#         return shape

#     def add_connector(self, s1, s2, **kwargs):
#         connector = add_connector(s1, s2, **kwargs)
#         connector.Name = 'gantt_connector'
#         return connector


# if __name__ == '__main__':
#     pp = xv.PowerPoint()
#     slide = pp.slides(1)
#     start = datetime.datetime(2018, 7, 1)
#     end = datetime.datetime(2018, 7, 30)
#     gc = GanttChart(start, end, 'daily')
#     gc.set_slide(slide)
#     gc.frame
#     gc.end

#     s1 = gc.add_point(gc.start, 0.5, 20, fill_color=rgb(255, 230, 210),
#                       line_color=rgb(0, 0, 0), line_weight=1)
#     s2 = gc.add_point(gc.end, 0.5, 20, fill_color=rgb(255, 230, 210),
#                       line_color=rgb(0, 0, 0), line_weight=1)

#     connector = gc.add_connector(s1, s2, weight=3, direction='vertical',
#                                  begin_arrow=True)

#     gc.add_line(datetime.datetime(2018, 5, 7),
#                 datetime.datetime(2023, 9, 30), 0.5)

#     s1.api.Name

#     slide.shapes.api.Count
#     shape = slide.shapes(3)
#     gc.to_date(shape.left)

#     connector.ConnectorFormat.BeginConnectedShape.Name
#     connector.Type
#     constants.msoConnectorStraingt
