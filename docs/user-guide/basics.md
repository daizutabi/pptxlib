# Basics

This page summarizes the core concepts and minimal API usage in pptxlib.

- Presentation: whole presentation (save/load, metadata)
- Slide: slides (add, layout, reorder)
- Shape: shapes (text, position/size, fill/line)
- Table: tables (cell edit, merge, width/height)

## Presentation and Slide

```python
from pptxlib import Presentation

p = Presentation()
slide = p.add_slide()
slide.add_title("タイトル")
p.save("basic.pptx")
```

## Shape basics

```python
from pptxlib import Presentation

with Presentation() as p:
    s = p.add_slide()
    r = s.add_shape("rectangle", 50, 80, 300, 80)
    r.text = "テキスト"
    p.save("shape-basic.pptx")
```

## Table basics

```python
from pptxlib import Presentation

with Presentation() as p:
    s = p.add_slide()
    t = s.add_table(2, 2, 50, 200, 300, 100)
    t.cell(0, 0).text = "A1"
    p.save("table-basic.pptx")
```
