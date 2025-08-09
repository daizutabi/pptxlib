# Styling (Fonts & Colors)

This page shows how to customize appearance focusing on fonts and colors.

```python exec="1" source="material-block"
from pptxlib import App
app = App()
pr = app.presentations.add()
slide = pr.slides.add(layout="TitleOnly")
slide.title = "Styled"
shape = slide.shapes.add("Rectangle", 100, 100, 240, 80)
shape.font.name = "Meiryo"
shape.font.size = 18
shape.font.bold = True
shape.font.color = "red"
shape.fill.color = (230, 240, 255)
app.quit()
```

TIPS:

- In addition to `Color.rgb(r, g, b)`, you can use `Color.theme(name)`
- `font` and `fill` provide common interfaces across many elements
