# Quick Start Guide

```python .md#_
from pptxlib import App
App().presentations.close()
```

This guide will help you get started with pptxlib quickly.
We'll create a simple presentation with a title slide
and a content slide.

## Creating a New App

You can create a new instance of the PowerPoint application
by calling the [`App`][pptxlib.app.App] class.

```python exec="1" source="material-block"
from pptxlib import App

app = App()
app
```

[`App.presentations`][pptxlib.presentation.Presentations]
is a collection of presentations.

```python exec="1" source="material-block"
app.presentations
```

## Creating a New Presentation

You can create a new presentation by calling the
[`add`][pptxlib.presentation.Presentations.add] method
on the [`presentations`][pptxlib.presentation.Presentations] collection.

```python exec="1" source="material-block"
pr = app.presentations.add()
pr
```

[`presentations`][pptxlib.presentation.Presentations]
attribute can be indexed by an integer
to access a specific presentation.

```python exec="1" source="material-block"
app.presentations[0]
```

## Adding a Title Slide

A title slide is a slide with a title and a subtitle optionally.

You can add a title slide by calling the
[`add`][pptxlib.slide.Slides.add] method
on the [`slides`][pptxlib.slide.Slides] collection
and passing the `layout` parameter
with the value `"Title"`.
Then, you can set the title of the slide by setting the
[`title`][pptxlib.slide.Slide.title] attribute.

```python exec="1" source="material-block"
slide = pr.slides.add(layout="Title")
slide.title = "Welcome to pptxlib"
```

Now, the [`slides`][pptxlib.slide.Slides]
collection has one slide.

```python exec="1" source="material-block"
pr.slides
```

Check the title of the slide.

```python exec="1" source="material-block"
pr.slides[0].title
```

## Adding Content Slides

You can add a content slide by calling the
[`add`][pptxlib.slide.Slides.add] method
on the [`slides`][pptxlib.slide.Slides] collection
and passing the `layout` parameter
with the layout name, for example, `"TitleOnly"`.

```python exec="1" source="material-block"
slide = pr.slides.add(layout="TitleOnly")
slide.title = "First Slide"
```

If you omit the `layout` parameter,
the layout of the previous slide is used.

```python exec="1" source="material-block"
slide = pr.slides.add()
slide.title = "Second Slide"
```

Now, we have three slides.

```python exec="1" source="material-block"
pr.slides
```

## Selecting a Slide

You can select a slide by calling the
[`select`][pptxlib.base.Element.select] method
on the slide object.

```python exec="1" source="material-block"
slide.select()
```

[`unselect`][pptxlib.app.App.unselect] method is also available.

```python exec="1" source="material-block"
app.unselect()
```

## Working with Shapes

You can add a shape to a slide by calling the
[`add`][pptxlib.shape.Shapes.add] method
on the [`shapes`][pptxlib.shape.Shapes] collection.

```python exec="1" source="material-block"
shape = slide.shapes.add("Rectangle", 100, 100, 200, 100)
shape
```

## Quit the App

```python exec="1" source="material-block"
app.quit()
```
