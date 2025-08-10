# Shapes

Working with shapes such as rectangles, lines, and images.

## Add and get

```python exec="1" source="material-block"
from pptxlib import App
app = App()
pr = app.presentations.add()
slide = pr.slides.add(layout="Blank")
rect = slide.shapes.add("Rectangle", 100, 100, 200, 100)
rect
app.quit()
```

## Tips

- Positions and sizes are specified in points (pt), not pixels
- You can search and reuse by name or type
