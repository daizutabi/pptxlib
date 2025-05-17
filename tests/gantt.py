from __future__ import annotations

from datetime import datetime

from pptxlib.core.app import App
from pptxlib.gantt import GanttFrame


def main():
    app = App()
    app.presentations.close()
    pr = app.presentations.add().size(400, 300)
    layout = pr.layouts.add("abc")
    slide = pr.slides.add(layout="TwoColumnText")
    pr.layouts.copy_from(slide, "def")
    # layouts = pr.api.SlideMaster.CustomLayouts
    # layout = pr.api.SlideMaster.CustomLayouts.Add(layouts.Count + 1)
    # layout.Name = "abc"
    for layout in pr.layouts:
        print(layout.name)
    # print(pr.slides.add().api.CustomLayout)

    # gantt = GanttFrame("day", datetime(2025, 4, 1), datetime(2025, 4, 10))

    # slide = pr.slides.add()
    # layout = slide.set_layout("GanttChart")
    # layout = slide.set_layout("GanttChart")
    # layout.shapes.add_table(2, 3, 100, 250, 100, 100)
    # print(layout.name)
    # for shape in layout.shapes:
    #     print(shape.name)


if __name__ == "__main__":
    main()
