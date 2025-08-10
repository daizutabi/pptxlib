# Tables

Adding tables and working with cells.

## テーブルの作成

```python exec="1" source="material-block"
from pptxlib import App
app = App()
pr = app.presentations.add()
slide = pr.slides.add(layout="Blank")
# Add a 3x2 table (x, y, width, height)
# Units are points (pt). Adjust as needed.

tbl = slide.shapes.add_table(3, 2, 100, 200, 400, 120)
# Set cell values (3 rows x 2 columns)
for r in range(3):
    for c in range(2):
        tbl[r, c].text = f"R{r+1}C{c+1}"
app.quit()
```

## Rows and columns

- Add/remove rows and columns and merge cells via `tbl.rows`, `tbl.columns`, and `tbl.merge`
