from win32com.client import CoClassBaseClass, DispatchBaseClass

from pptxlib.app import (
    PowerPoint,
    Presentation,
    Presentations,
    Slide,
    Slides,
)
from pptxlib.core import (
    Base,
    Collection,
    Element,
)


def test_powerpoint(pp: PowerPoint):
    assert isinstance(pp, Base)
    assert isinstance(pp, PowerPoint)
    assert isinstance(pp.api, DispatchBaseClass)
    assert pp.name == "Microsoft PowerPoint"


def test_presentations(prs: Presentations):
    assert isinstance(prs, Collection)
    assert isinstance(prs, Presentations)
    assert isinstance(prs.api, DispatchBaseClass)
    assert prs.name == "Presentations"
    assert len(prs) == 0


def test_presentations_add_close(prs: Presentations):
    pr = prs.add()
    assert isinstance(pr, Element)
    assert isinstance(pr, Presentation)
    assert isinstance(pr.api, CoClassBaseClass)
    assert len(prs) == 1
    assert pr.name[-1].isdigit()
    pr.close()
    assert len(prs) == 0


def test_presentations_active(prs: Presentations):
    pr1 = prs.add()
    pr2 = prs.add()
    assert len(prs) == 2
    assert prs.active.name == pr2.name
    pr1.close()
    pr2.close()
    assert len(prs) == 0


def test_presentations_call(prs: Presentations, pr_list: list[Presentation]):
    for i in range(3):
        assert prs(i + 1).name == pr_list[i].name
    assert prs().name == pr_list[2].name


def test_presentations_iter(prs: Presentations, pr_list: list[Presentation]):
    assert len(list(prs)) == 3
    names = [pr.name for pr in pr_list]
    for k, pr in enumerate(prs):
        assert pr.name == names[k]


def test_presentations_getitem(prs: Presentations, pr_list: list[Presentation]):
    for i in range(3):
        pr = prs[i]
        assert isinstance(pr, Presentation)
        assert pr.name == pr_list[i].name


def test_presentations_getitem_slice(prs: Presentations, pr_list: list[Presentation]):
    pr = prs[0:2]
    assert isinstance(pr, list)
    assert len(pr) == 2
    assert pr[-1].name == pr_list[1].name


def test_presentations_getitem_neg(prs: Presentations, pr_list: list[Presentation]):
    pr = prs[-2]
    assert isinstance(pr, Presentation)
    assert pr.name == pr_list[1].name


def test_presentation_size(pr: Presentation):
    assert pr.width > 100
    assert pr.height > 100


def test_slides(slides: Slides):
    assert isinstance(slides, Collection)
    assert isinstance(slides, Slides)
    assert isinstance(slides.api, DispatchBaseClass)
    assert slides.name == "Slides"
    assert len(slides) == 0


def test_slides_add_delete(slides: Slides):
    slide = slides.add()
    assert isinstance(slide, Element)
    assert isinstance(slide, Slide)
    assert isinstance(slide.api, CoClassBaseClass)
    assert len(slides) == 1
    assert slide.name == "Slide1"
    slide.delete()
    assert len(slides) == 0


def test_slides_active_select(slides: Slides):
    s1 = slides.add()
    s2 = slides.add()
    assert len(slides) == 2
    assert slides.active.name == s1.name
    s2.select()
    assert slides.active.name == s2.name
    s1.delete()
    s2.delete()
    assert len(slides) == 0


def test_slide_title(slide: Slide):
    assert slide.title == ""
    slide.title = "Slide Title"
    assert slide.title == "Slide Title"


def test_repr_powerpoint(pp: PowerPoint):
    assert repr(pp) == "<PowerPoint>"


def test_repr_presentations(prs: Presentations):
    assert repr(prs) == "<Presentations>"


def test_repr_presentation(pr: Presentation):
    assert repr(pr).startswith("<Presentation [")


def test_repr_slides(slides: Slides):
    assert repr(slides) == "<Slides>"


def test_repr_slide(slide: Slide):
    assert repr(slide) == "<Slide [Slide1]>"
