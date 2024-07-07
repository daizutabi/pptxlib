# import numpy as np
# import pandas as pd

from xlviews.decorators import wait_updating
from xlviews.frame import SheetFrame
# import xlviews as xv
from xlviews.style import rgb, set_alignment, set_border, set_fill, set_font


class MultiIndexFrame(object):
    @wait_updating
    def __init__(self, *args, data=None, **kwargs):
        """
        をプロットするためにシートフレームを作成する．
        """
        self.data = data
        self.columns = []
        self.sf_index = None
        self.sf = {}
        sf = None
        for column in self.data.columns.get_level_values(0):
            if column not in self.columns:
                if not self.columns:
                    self.sf_index = SheetFrame(*args, data=self.data[column],
                                               index=True, **kwargs)
                    col = (self.sf_index.column +
                           len(self.sf_index.index_columns))
                else:
                    col = sf.column + len(sf.columns)
                max_columns = len(self.data[column].columns)
                sf = SheetFrame(self.sf_index.sheet, self.sf_index.row, col,
                                data=self.data[column], index=False,
                                max_columns=max_columns, **kwargs)
                self.sf[column] = sf
                self.columns.append(column)
        self.set_style()

    def set_style(self):
        for column, sf in self.sf.items():
            start = sf.cell.offset(-1, 0)
            end = start.offset(0, len(sf.columns) - 1)
            range_ = sf.sheet.range(start, end)
            range_.api.Merge()
            range_.value = column
            set_font(range_, bold=True, color='white', size=9)
            set_alignment(range_, horizontal_alignment='center')
            set_border(range_, edge_width=3)
            set_fill(range_, color=rgb(0, 80, 180))


def main():
    import mtj
    import xlwings as xw

    xw.apps.add()
    book = xw.books[0]
    sheet = book.sheets[0]

    directory = mtj.get_directory('local', 'Data')
    run = mtj.get_paths_dataframe(directory, 'SL1048-01', tool='MRP',
                                  recipe='HR', name='28shots')
    path = mtj.get_path(directory, run.loc[0])

    def describe(df):
        df = df.describe()
        df.rename(index={'50%': 'median'}, inplace=True)
        soa = df.loc['std'] / df.loc['median']
        soa.name = 'soa'
        df = df.append(soa)
        return df.loc[['median', 'soa', 'std', 'count', 'min', '25%',
                       '75%', 'max']]

    with mtj.data(path) as data:
        data.merge_device()
        columns = ['Rmin', 'Rmax', 'TMR']
        index = ['wafer', 'cad', 'sx', 'sy']
        df = data[index + columns]
        df = df.groupby(index)[columns].apply(describe).unstack()
        sf = SheetFrame(sheet, 2, 2, data=df)
        # MultiIndexFrame(sheet, 3, 2, data=df)
        print(sf.get_adjacent_cell())
        print(sf.columns)
        print(sf[('Rmin', 'soa')])

# Sub Macro1()
# '
# ' Macro1 Macro
# '
#
# '
#     Columns("V:AB").Select
#     Selection.Columns.Group
#     Columns("V:AA").Select
#     Selection.Columns.Group
# End Sub
# Sub Macro2()
# '
# ' Macro2 Macro
# '
#
# '
#     With ActiveSheet.Outline
#         .AutomaticStyles = False
#         .SummaryRow = xlBelow
#         .SummaryColumn = xlLeft
#     End With
# End Sub


if __name__ == '__main__':
    main()
