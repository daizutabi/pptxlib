import pandas as pd
import xlwings as xw

from xlviews.decorators import wait_updating
from xlviews.element import Bar
from xlviews.frame import SheetFrame
from xlviews.style import set_alignment, set_font, set_number_format
from xlviews.utils import columns_list, rgb


class BoxFrame(SheetFrame):
    values = ['min', '25%', '50%', '75%', 'max']
    tail_values = ['be', 'bt', 'bb', 'tb', 'te']

    @wait_updating
    def __init__(self, *args, data=None, columns=None, by=None,
                 number_format=None, style=True, autofit=True, **kwargs):
        """
        分布をプロットするためにシートフレームを作成する．
        Parameters
        ----------
        dist : str or dict
        """
        by = columns_list(data, by)
        if columns is None:
            columns = [column for column in data.columns
                       if column not in by]
        grouped = data.groupby(by)
        df = grouped[columns].describe()
        dfs = []
        for column in columns:
            dfs.append(df[column][BoxFrame.values])
        df = pd.concat(dfs, axis=1)  # type: pd.DataFrame

        super().__init__(*args, data=df, style=False, **kwargs)

        cell = self.cell.offset(len(self) + 3)
        self.tail = SheetFrame(cell, data=df, style=False)
        tail_columns = BoxFrame.tail_values * len(columns)
        self.tail.cell.offset(0, self.tail.index_level).value = tail_columns

        for sf in [self, self.tail]:
            add_wide_column(sf, columns)
            if number_format:
                for key, value in number_format.items():
                    set_number_format(sf.range(key, -1), value)

        for column in columns:
            refs = list(get_references(self, column))
            formulas = [f'{refs[1]}-{refs[0]}', f'{refs[1]}',
                        f'{refs[2]}-{refs[1]}', f'{refs[3]}-{refs[2]}',
                        f'{refs[4]}-{refs[3]}']
            set_references(self, column, formulas)

        if style:
            self.set_style()
            self.tail.set_style(gray=True)
        if autofit:
            self.autofit()

    @wait_updating
    def plot(self, x, y, **kwargs):
        elem = Bar(x, [(y, 'bt'), (y, 'bb'), (y, 'tb')], data=self.tail,
                   stacked=True, label=None, **kwargs)
        k = len(elem.series_collection) - 3
        bottom = elem.series_collection[k]
        bottom.Format.Fill.Visible = False
        bottom.HasErrorBars = True
        top = elem.series_collection[k + 2]
        top.HasErrorBars = True

        direction = xw.constants.ErrorBarDirection.xlY
        type_ = xw.constants.ErrorBarType.xlErrorBarTypeCustom

        # AmountとMinusValuesの両方が必要
        be = self.tail.range((y, 'be'), -1).api
        include = xw.constants.ErrorBarInclude.xlErrorBarIncludeMinusValues
        bottom.ErrorBar(Direction=direction, Include=include, Type=type_,
                        Amount=be, MinusValues=be)
        beb = bottom.ErrorBars
        bottom = elem.series_collection[k + 1]
        te = self.tail.range((y, 'te'), -1).api
        include = xw.constants.ErrorBarInclude.xlErrorBarIncludePlusValues
        top.ErrorBar(Direction=direction, Include=include, Type=type_,
                     Amount=te, MinusValues=te)
        teb = top.ErrorBars

        for obj in [teb, beb]:
            line = obj.Format.Line
            line.ForeColor.RGB = rgb(0, 0, 100)
            line.Weight = 1.5
        for obj in [top, bottom]:
            line = obj.Format.Line
            line.ForeColor.RGB = rgb(0, 0, 100)
            line.Weight = 1
            fill = obj.Format.Fill
            fill.ForeColor.RGB = rgb(0, 0, 100)
            fill.Transparency = 0.8


def get_references(sf, column):
    for value in BoxFrame.values:
        if value.endswith('%'):
            value = float(value[:-1]) / 100
        yield sf.range((column, value)).get_address(row_absolute=False)


def set_references(sf, column, formulas):
    for value, formula in zip(BoxFrame.tail_values, formulas):
        sf.tail.range((column, value), -1).value = '=' + formula


def add_wide_column(sf, columns):
    cell = sf.cell.offset(-1, sf.index_level)
    for column in columns:
        cell.value = column
        set_font(cell, bold=True, color='#002255')
        set_alignment(cell, horizontal_alignment='left')
        cell = cell.offset(0, len(BoxFrame.values))


def main():
    import mtj
    directory = mtj.get_directory('local', 'Data')
    run = mtj.get_paths_dataframe(directory, 'SL1050-01', recipe='HR')
    series = run.iloc[0]
    path = mtj.get_path(directory, series)
    with mtj.data(path, device=True) as data:
        df = data[['wafer', 'cad', 'Rmin', 'Rmax']]
        bf = BoxFrame(3, 2, data=df, by=':cad',
                      number_format={'Rmin': '0.00', 'Rmax': '0.00'})
        bf.plot('wafer', 'Rmin', yticks=(0, 10, 1))


if __name__ == '__main__':
    main()
