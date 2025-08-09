# Slides & Layouts

Tips to work with slides efficiently.

## Add and select slides

```python exec="1" source="material-block"
from pptxlib import App
app = App()
pr = app.presentations.add()
slide1 = pr.slides.add(layout="Title")
slide1.title = "タイトル"
slide2 = pr.slides.add()  # 直前のレイアウトを継承
slide2.select()
app.unselect()
app.quit()
```

## Layouts

- Supports presets such as "Title", "TitleOnly", and "Blank"
- If you omit the layout, the previous slide’s layout is inherited
