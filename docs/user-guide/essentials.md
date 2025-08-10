# Essentials

Let’s go through the basic flow of pptxlib. Here’s a minimal example that starts the app, creates a presentation, adds a slide, and quits.

```python exec="1" source="material-block"
from pptxlib import App

app = App()
pr = app.presentations.add()
slide = pr.slides.add(layout="TitleOnly")
slide.title = "Hello pptxlib"
app.quit()
```

Notes:

- App represents the PowerPoint COM application
- Create a new presentation with `presentations.add()`
- Add a slide with `slides.add()` (if you omit layout, it inherits the previous slide’s layout)
