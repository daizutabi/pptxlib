"""
ショットマップを作成する．
"""
import numpy as np
from pandas import DataFrame
import xlwings as xw
from xlviews.decorators import wait_updating
from xlviews.style import (
    set_border,
    set_font,
    set_fill,
    set_alignment,
    set_number_format,
)
from xlviews.utils import get_sheet_cell_row_column, rgb, format_label
from spin.pandas import select
from spin.pandas.dataset import DataSet
import spin.pandas as spd


class ShotMap(object):
    @wait_updating
    def __init__(
        self,
        *args,
        data=None,
        value=None,
        primary_x="sx",
        primary_y="sy",
        secondary_x="dx",
        secondary_y="dy",
        primary_xvalues=None,
        primary_yvalues=None,
        font_size=7,
        column_width=2,
        row_height=10,
        vmin=None,
        vmax=None,
        inside_width=0,
        edge_width=2,
        title=None,
        name="auto",
        sel=None,
        edge_color=rgb(100, 100, 100),
        inside_color=rgb(240, 240, 240),
        empty_shot_color=rgb(230, 230, 230),
        invalid_shot_color=rgb(190, 190, 190),
        valid_shot=None,
        callback=None,
        agg=None,
        link=True,
        unit=None,
    ):
        self.sheet, self.cell, row, column = get_sheet_cell_row_column(*args)
        self.shots = {}

        cell_vmin = self.cell
        cell_vmin.value = vmin
        set_font(cell_vmin, bold=True, color="blue")
        cell_vmax = cell_vmin.offset(0, 1)
        cell_vmax.value = vmax
        set_font(cell_vmax, bold=True, color="red")

        vmin_ = cell_vmin.get_address()
        vmax_ = cell_vmax.get_address()

        from xlviews.frame import SheetFrame

        if isinstance(data, SheetFrame):
            number_format = data.get_number_format(value)
        else:
            number_format = None

        if agg is not None:
            if not isinstance(data, SheetFrame):
                raise ValueError("funcが指定されたときはSheetFrameのみ可能")
            df = data.data.reset_index()
            if sel:
                df = df[data.select(**sel)]
            const = spd.const_values(df)

            include_sheetname = self.sheet != data.sheet
            data = data.aggregate(
                agg,
                value,
                by=[primary_x, primary_y],
                sel=sel,
                include_sheetname=include_sheetname,
            )
            data["formula"] = data["formula"].map(lambda formula: f"={formula}")
            for column_ in const:
                data[column_] = const[column_].iloc[0]
            data[secondary_x] = 1
            data[secondary_y] = 1
            value = "formula"
            sel = None

        if isinstance(data, DataFrame):
            df = select(data, sel).copy()
        elif isinstance(data, SheetFrame):
            include_sheetname = self.sheet != data.sheet
            df = data.data.reset_index()
            if link:
                values = data.get_address(
                    value, include_sheetname=include_sheetname, formula=True
                )
                df[value] = values
            df = select(df, sel).copy()
        elif isinstance(data, DataSet):
            columns = data.required_columns(fstring=title)
            sel_ = sel if sel else {}
            df = data.get(
                [primary_x, primary_y, secondary_x, secondary_y, value, *columns],
                **sel_,
            )
        else:
            raise ValueError("不明な型:", type(data))

        if title:
            if isinstance(title, str) or callable(title):
                title = format_label(df, title)
            cell_title = cell_vmax.offset(0, 1)
            cell_title.value = title
            set_font(cell_title, bold=True)
            self.cell = self.cell.offset(1)
            row += 1

        def unique_index(array):
            index = array.sort_values().unique()
            dict_ = {index: key for key, index in enumerate(index)}
            return array.map(dict_)

        if callback:
            df = df.groupby([primary_x, primary_y]).apply(callback)
        df[secondary_x] = unique_index(df[secondary_x])
        df[secondary_y] = unique_index(df[secondary_y])

        dx_len = int(df[secondary_x].max() + 1)
        dy_len = int(df[secondary_y].max() + 1)

        if primary_xvalues is None and valid_shot is not None:
            primary_xvalues = sorted(valid_shot[primary_x].unique())
            primary_xvalues = [int(x) for x in primary_xvalues]
        elif isinstance(primary_xvalues, tuple):
            primary_xvalues = range(primary_xvalues[0], primary_xvalues[1] + 1)
        if primary_yvalues is None and valid_shot is not None:
            primary_yvalues = sorted(valid_shot[primary_y].unique())
            primary_yvalues = [int(y) for y in primary_yvalues]
        elif isinstance(primary_yvalues, tuple):
            primary_yvalues = range(primary_yvalues[0], primary_yvalues[1] + 1)
        primary_xvalues = list(primary_xvalues)
        primary_yvalues = list(primary_yvalues)

        # for shot label
        row += 1
        column += 1
        self.cell = self.cell.offset(1, 1)

        end = row, column
        for sx in primary_xvalues:
            for sy in primary_yvalues:
                start = (
                    row + (sy - primary_yvalues[0]) * dy_len,
                    column + (sx - primary_xvalues[0]) * dx_len,
                )
                end = start[0] + dy_len - 1, start[1] + dx_len - 1
                range_ = self.sheet.range(start, end)
                df_ = df[(df[primary_x] == sx) & (df[primary_y] == sy)]
                if len(df_):
                    # print(df_)
                    df_ = df_.pivot(
                        index=secondary_y, columns=secondary_x, values=value
                    )
                    self.sheet.range(start).value = df_.values
                    set_alignment(range_, horizontal_alignment="center")
                    if number_format:
                        set_number_format(range_, number_format)
                set_border(
                    range_,
                    inside_width=inside_width,
                    edge_width=edge_width,
                    inside_color=inside_color,
                    edge_color=edge_color,
                )

                values = range_.options(np.array).value
                if all(np.isnan(values.flatten())):
                    range_.api.Interior.Color = empty_shot_color
                if valid_shot is not None:
                    query = f"{primary_x} == @sx and {primary_y} == @sy"
                    if len(valid_shot.query(query)) == 0:
                        range_.api.Interior.Color = invalid_shot_color

                self.shots[(sx, sy)] = range_

                if sx == primary_xvalues[0]:
                    row_, column_ = start
                    range_ = self.sheet.range(
                        (row_, column_ - 1), (row_ + dy_len - 1, column_ - 1)
                    )
                    set_shot_label(range_, sy, edge_width=edge_width)
                if sy == primary_yvalues[0]:
                    row_, column_ = start
                    range_ = self.sheet.range(
                        (row_ - 1, column_), (row_ - 1, column_ + dx_len - 1)
                    )
                    set_shot_label(range_, sx, edge_width=edge_width)

        range_ = self.sheet.range((row, column), end)
        set_border(range_, inside_width=0, edge_width=edge_width)
        set_font(range_, size=font_size, color=rgb(250, 250, 250))
        range_.column_width = column_width
        range_.row_height = row_height

        colors = [rgb(130, 130, 255), rgb(80, 185, 80), rgb(255, 130, 130)]
        values = [f"={vmin_}", f"=({vmin_} + {vmax_}) / 2", f"={vmax_}"]
        set_color_condition(range_, values, colors)

        self.start = self.cell
        # shot label column width and row height
        self.start.offset(0, -1).column_width = 2.3
        self.start.offset(-1, 0).row_height = row_height
        self.end = self.sheet.range(*end)
        self.end.offset(0, 1).column_width = 1
        self.range = range_
        self.width = end[1] - self.start.column + 2  # for label
        self.height = end[0] - self.start.row + 2  # for label

        self.end.offset(0, 2).column_width = 4
        self.end.offset(0, 3).column_width = 2
        color_bar_column = self.end.offset(0, 2).column
        range_ = self.sheet.range(
            (self.start.row, color_bar_column), (end[0], color_bar_column)
        )
        n = len(range_) - 1
        color_bar_values = [
            f"=({n}-{k})/{n}*{vmin_}+{k}/{n}*{vmax_}" for k in range(n, -1, -1)
        ]
        range_.options(transpose=True).value = color_bar_values
        set_color_condition(range_, values, colors)
        set_border(range_, inside_width=0, edge_width=edge_width)
        set_font(range_, size=1, color=rgb(255, 255, 255))
        set_alignment(range_, horizontal_alignment="center")
        set_font(range_[0], size=9, color="white", bold=True)
        set_font(range_[-1], size=9, color="white", bold=True)
        if unit:
            unit_cell = range_[0].offset(-1)
            unit_cell.value = unit
            set_alignment(unit_cell, horizontal_alignment="center")
            set_font(unit_cell, size=9, bold=True)

        self.width += 2

        if name == "auto" and title:
            name = title
        if isinstance(name, str) and name != "auto":
            refers_to = "=" + self.range.get_address()
            name = name.replace("-", "__")
            self.sheet.names.add(name, refers_to)


def set_shot_label(range_, label, edge_width=2):
    range_.api.Merge()
    range_.value = label
    set_font(range_, bold=True)
    set_border(range_, edge_width=edge_width, edge_color="black")
    set_fill(range_, rgb(255, 255, 200))
    set_alignment(range_, horizontal_alignment="center", vertical_alignment="center")


def set_color_condition(range_, values, colors):
    condition = range_.api.FormatConditions.AddColorScale(3)
    condition.SetFirstPriority()
    type_ = xw.constants.ConditionValueTypes.xlConditionValueNumber
    for k in range(3):
        criteria = condition.ColorScaleCriteria(k + 1)
        criteria.Type = type_
        criteria.Value = values[k]
        criteria.FormatColor.Color = colors[k]


def main():
    import mtj

    process = "S6579"
    directory = mtj.get_directory("local", "Data")
    path = mtj.get_path(
        directory, process, "SL1048-05", "MRP", "HR", "HR_32pad_28shots"
    )
    data = mtj.data(path, device=True)

    xw.apps.add()
    book = xw.books(1)
    sheet = book.sheets(1)

    from mtj.data.common import get_shot

    shot = get_shot(process, stack=True)

    def callback(df):
        div = 6
        df["dx"] = (df["dy"] - 1) // div + 1
        df["dy"] = (df["dy"] - 1) % div + 1
        return df

    kwargs = dict(
        value="Rmin",
        title="{wafer}_{cad}",
        vmin=0,
        vmax=10,
        callback=callback,
        valid_shot=shot,
    )

    book.app.api.ActiveWindow.Zoom = 40

    # DataFrame
    # df = data[['wafer', 'cad', 'sx', 'sy', 'dx', 'dy', 'Rmin']]
    # shotmap = ShotMap(sheet, 2, 2, data=df, sel={'cad': 'Y60'},
    #                   **kwargs)

    # DataSet
    # column = shotmap.cell.column + shotmap.width
    # shotmap = ShotMap(sheet, 2, column, data=data, sel={'cad': 'Y70'},
    #                   **kwargs)

    # SheetFrame
    from xlviews.frame import SheetFrame

    sheet = book.sheets.add()
    # book.app.api.ActiveWindow.Zoom = 40
    sf = SheetFrame(
        sheet,
        2,
        2,
        data=data,
        columns=["wafer", "cad", "sx", "sy", "dx", "dy", "Rmin"],
        index=":sy",
        sort_index=True,
    )
    sheet = book.sheets.add()
    cell = sheet.range(2, 2)
    ShotMap(cell, data=sf, **kwargs, sel={"cad": "Y70"}, unit="kΩ")
    # shotmap = ShotMap(cell, data=sf, sel={'cad': 'Y60'}, func='median',
    #                   **kwargs)
    # grid = xv.FacetGrid(data=sf, x='cad', cell=cell.offset(14))
    # grid.map(ShotMap, func='median', **kwargs)

    # Grid + DataSet
    # sheet = book.sheets.add()
    # book.app.api.ActiveWindow.Zoom = 40
    # grid = xv.FacetGrid(data=data, x='cad', sheet=sheet, row=2, column=2)
    # shotmaps = grid.map(ShotMap, **kwargs)


if __name__ == "__main__":
    main()
