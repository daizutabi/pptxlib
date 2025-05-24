# API Reference Overview

The pptxlib API is organized into several main components that correspond to the
PowerPoint object model. This section provides an overview of the main classes
and their relationships.

## Core Components

### App

The `App` class is the entry point to the PowerPoint application. It provides
access to presentations and manages the PowerPoint application instance.

```python
from pptxlib import App

with App() as app:
    # Work with presentations
    presentations = app.presentations
```

### Presentations

The `Presentations` class represents a collection of presentations. It provides
methods to create new presentations and access existing ones.

```python
from pptxlib import App

with App() as app:
    # Create a new presentation
    presentation = app.presentations.add()

    # Open an existing presentation
    existing = app.presentations.open("path/to/presentation.pptx")
```

### Slides

The `Slides` class represents a collection of slides within a presentation. It
provides methods to add, remove, and modify slides.

```python
from pptxlib import App

with App() as app:
    presentation = app.presentations.add()

    # Add a new slide
    slide = presentation.slides.add()

    # Access existing slides
    first_slide = presentation.slides[0]
```

### Shapes

The `Shapes` class represents a collection of shapes on a slide. It provides
methods to add and modify various types of shapes.

```python
from pptxlib import App

with App() as app:
    presentation = app.presentations.add()
    slide = presentation.slides.add()

    # Add different types of shapes
    textbox = slide.shapes.add_textbox("Hello")
    rectangle = slide.shapes.add_shape("rectangle", 100, 100, 200, 100)
    table = slide.shapes.add_table(3, 3, 100, 100, 200, 100)
```

## Utility Classes

### Color

The `Color` class provides a way to specify colors in RGB format.

```python
from pptxlib import App, Color

with App() as app:
    presentation = app.presentations.add()
    slide = presentation.slides.add()

    # Create a red text box
    textbox = slide.shapes.add_textbox("Red Text")
    textbox.text_frame.font.color = Color(255, 0, 0)
```

### Font

The `Font` class provides properties for text formatting.

```python
from pptxlib import App

with App() as app:
    presentation = app.presentations.add()
    slide = presentation.slides.add()

    # Format text
    textbox = slide.shapes.add_textbox("Formatted Text")
    textbox.text_frame.font.size = 24
    textbox.text_frame.font.bold = True
```

## Detailed API Documentation

For detailed information about each class and its methods, please refer to the
following pages:

- [pptxlib Module](pptxlib.md)
- [App Class](pptxlib.md#app)
- [Presentations Class](pptxlib.md#presentations)
- [Slides Class](pptxlib.md#slides)
- [Shapes Class](pptxlib.md#shapes)
- [Color Class](pptxlib.md#color)
- [Font Class](pptxlib.md#font)
