# Quick Tour

A five-minute tour to experience pptxlib quickly.

## 0. Prerequisites

- Windows with Microsoft PowerPoint installed
- Python 3.11+

## 1. Create a PowerPoint application instance

```python exec="1" source="1"
from pptxlib import App
app = App()
app
```

## 1. Create a single-slide presentation with the smallest script

```python
from pptxlib import Presentation

with Presentation() as p:
    slide = p.add_slide()
    slide.add_title("Hello, pptxlib!")
    p.save("hello.pptx")
```

Open `hello.pptx` in PowerPoint and you’ll see one slide.

## 2. Add a shape and a table

```python
from pptxlib import Presentation

with Presentation() as p:
    slide = p.add_slide()
    rect = slide.add_shape("rectangle", left=50, top=80, width=300, height=80)
    rect.text = "四角形"

    table = slide.add_table(3, 3, left=50, top=200, width=400, height=120)
    table.cell(0, 0).text = "A1"
    p.save("tour.pptx")
```

## 3. Fonts and colors

```python
from pptxlib import Presentation, Font, Color

with Presentation() as p:
    s = p.add_slide()
    t = s.add_title("スタイル例")
    t.font = Font(name="Segoe UI", size=28, bold=True, color=Color.rgb(0, 90, 158))
    p.save("styled.pptx")
```
