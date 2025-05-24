# Quick Start Guide

This guide will help you get started with pptxlib quickly. We'll create a simple
presentation with a title slide and a content slide.

## Creating a New Presentation

```python
from pptxlib import App

# Create a new presentation
with App() as app:
    # Add a new presentation
    presentation = app.presentations.add()

    # Add a title slide
    title_slide = presentation.slides.add()
    title_shape = title_slide.shapes.add_textbox("Welcome to pptxlib")
    title_shape.text_frame.text = "Welcome to pptxlib"

    # Add a content slide
    content_slide = presentation.slides.add()
    content_shape = content_slide.shapes.add_textbox("Getting Started")
    content_shape.text_frame.text = "Getting Started"

    # Save the presentation
    presentation.save_as("quickstart.pptx")
```

## Working with Shapes

```python
from pptxlib import App

with App() as app:
    presentation = app.presentations.add()
    slide = presentation.slides.add()

    # Add a text box
    textbox = slide.shapes.add_textbox("Hello, World!")

    # Add a rectangle
    rectangle = slide.shapes.add_shape("rectangle", 100, 100, 200, 100)

    # Add a table
    table = slide.shapes.add_table(3, 3, 100, 100, 200, 100)
```

## Customizing Text and Colors

```python
from pptxlib import App, Color

with App() as app:
    presentation = app.presentations.add()
    slide = presentation.slides.add()

    # Add text with custom formatting
    textbox = slide.shapes.add_textbox("Formatted Text")
    textbox.text_frame.text = "Formatted Text"
    textbox.text_frame.font.size = 24
    textbox.text_frame.font.color = Color(255, 0, 0)  # Red color
```

## Next Steps

Now that you've learned the basics, you can explore more features:

- [Basic Usage](../user-guide/basic-usage.md)
- [Working with Presentations](../user-guide/presentations.md)
- [Managing Slides](../user-guide/slides.md)
- [Working with Shapes](../user-guide/shapes.md)
- [Creating Tables](../user-guide/tables.md)
