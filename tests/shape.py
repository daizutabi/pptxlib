from __future__ import annotations

from pathlib import Path

from pptxlib.core.app import App


def main():
    app = App()
    app.presentations.close()
    pr = app.presentations.add()
    slide = pr.slides.add()
    shapes = slide.shapes
    s1 = shapes.add("Rectangle", 100, 100, 100, 100)
    s2 = shapes.add("Oval", 150, 150, 90, 80)
    slide.export(Path(__file__).parent / "a.svg")


if __name__ == "__main__":
    main()
