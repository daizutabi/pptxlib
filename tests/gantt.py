from __future__ import annotations

from datetime import datetime

from pptxlib.core.app import App
from pptxlib.gantt import GanttChart


def main():
    app = App()
    app.presentations.close()
    pr = app.presentations.add()
    gc = GanttChart(
        "day",
        datetime(2025, 4, 1),
        datetime(2025, 4, 7),
        pr,
        20,
        150,
        bottom=10,
    )
    pr.slides.add(layout=gc.layout)
    print(gc.frame.columns)


if __name__ == "__main__":
    main()
