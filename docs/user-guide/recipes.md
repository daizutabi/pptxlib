# Recipes

Frequently used snippets organized by use case.

## Generate slides with sequential titles

```python
from pptxlib import Presentation

titles = [f"章 {i}" for i in range(1, 6)]
with Presentation() as p:
    for t in titles:
        s = p.add_slide()
        s.add_title(t)
    p.save("chapters.pptx")
```

## Apply a common style to shapes

```python
from pptxlib import Presentation, Font, Color

brand = dict(font=Font(name="Segoe UI", size=18), fill=Color.rgb(235, 245, 255))

with Presentation() as p:
    s = p.add_slide()
    box = s.add_shape("round-rectangle", 40, 80, 420, 120)
    box.text = "ブランドボックス"
    box.apply(**brand)
    p.save("brand.pptx")
```

## Create a table from a DataFrame (values only)

```python
import pandas as pd
from pptxlib import Presentation

with Presentation() as p:
    s = p.add_slide()
    df = pd.DataFrame({"A":[1,2],"B":[3,4]})
    t = s.add_table(df.shape[0]+1, df.shape[1], 40, 80, 420, 120)
    for j,col in enumerate(df.columns):
        t.cell(0,j).text = col
    for i,row in df.iterrows():
        for j,val in enumerate(row):
            t.cell(i+1,j).text = str(val)
    p.save("from-excel.pptx")
```
