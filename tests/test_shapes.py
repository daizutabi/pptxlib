from win32com.client import CoClassBaseClass, DispatchBaseClass

from pptxlib.app import Slide
from pptxlib.core import Collection, Element
from pptxlib.shapes import Shape, Shapes


def test_shapes(shapes: Shapes):
    assert isinstance(shapes, Collection)
    assert isinstance(shapes, Shapes)
    assert isinstance(shapes.api, DispatchBaseClass)
    assert shapes.name == "Shapes"
    assert len(shapes) == 1  # title shape


def test_title_shape(title_shape: Shape):
    shape = title_shape
    assert isinstance(shape, Element)
    assert isinstance(shape, Shape)
    assert isinstance(shape.api, DispatchBaseClass)
    assert shape.name == "Title 1"


def test_text_range(title_shape: Shape):
    text_range = title_shape.text_range
    assert isinstance(text_range, DispatchBaseClass)


def test_text(title_shape: Shape):
    assert title_shape.text == ""
    title_shape.text = "Title"
    assert title_shape.text == "Title"


# def test_shapes_add_delete(slide: Slide):
#     assert isinstance(prs, Collection)
#     assert isinstance(prs, Presentations)
#     assert isinstance(prs.api, DispatchBaseClass)
#     assert prs.name == "Presentations"
#     assert len(prs) == 0


# def test_presentations_add_close(prs: Presentations):
#     pr = prs.add()
#     assert isinstance(pr, Element)
#     assert isinstance(pr, Presentation)
#     assert isinstance(pr.api, CoClassBaseClass)
#     assert len(prs) == 1
#     assert pr.name[-1].isdigit()
#     pr.close()
#     assert len(prs) == 0


# def test_presentations_active(prs: Presentations):
#     pr1 = prs.add()
#     pr2 = prs.add()
#     assert len(prs) == 2
#     assert prs.active.name == pr2.name
#     pr1.close()
#     pr2.close()
#     assert len(prs) == 0


# def test_presentations_call(prs: Presentations, pr_list: list[Presentation]):
#     for i in range(3):
#         assert prs(i + 1).name == pr_list[i].name


# def test_presentations_iter(prs: Presentations, pr_list: list[Presentation]):
#     assert len(list(prs)) == 3
#     names = [pr.name for pr in pr_list]
#     for pr in prs:
#         assert pr.name in names


# def test_presentations_getitem(prs: Presentations, pr_list: list[Presentation]):
#     for i in range(3):
#         pr = prs[i]
#         assert isinstance(pr, Presentation)
#         assert pr.name == pr_list[i].name


# def test_presentations_getitem_slice(prs: Presentations, pr_list: list[Presentation]):
#     pr = prs[0:2]
#     assert isinstance(pr, list)
#     assert len(pr) == 2
#     assert pr[-1].name == pr_list[1].name


# def test_presentations_getitem_neg(prs: Presentations, pr_list: list[Presentation]):
#     pr = prs[-2]
#     assert isinstance(pr, Presentation)
#     assert pr.name == pr_list[1].name


# def test_slides(slides: Slides):
#     assert isinstance(slides, Collection)
#     assert isinstance(slides, Slides)
#     assert isinstance(slides.api, DispatchBaseClass)
#     assert slides.name == "Slides"
#     assert len(slides) == 0


# def test_slides_add_delete(slides: Slides):
#     slide = slides.add()
#     assert isinstance(slide, Element)
#     assert isinstance(slide, Slide)
#     assert isinstance(slide.api, CoClassBaseClass)
#     assert len(slides) == 1
#     assert slide.name == "Slide1"
#     slide.delete()
#     assert len(slides) == 0


# def test_slides_active_select(slides: Slides):
#     s1 = slides.add()
#     s2 = slides.add()
#     assert len(slides) == 2
#     assert slides.active.name == s1.name
#     s2.select()
#     assert slides.active.name == s2.name
#     s1.delete()
#     s2.delete()
#     assert len(slides) == 0


# def test_slide_title(slide: Slide):
#     print(slide.title)
#     assert 0


# # def test_presentations_call(prs: Presentations, pr_list: list[Presentation]):
# #     for i in range(3):
# #         assert prs(i + 1).name == pr_list[i].name


# def test_repr_powerpoint(pp: PowerPoint):
#     assert repr(pp) == "<PowerPoint>"


# def test_repr_presentations(prs: Presentations):
#     assert repr(prs) == "<Presentations>"


# def test_repr_presentation(pr: Presentation):
#     assert repr(pr).startswith("<Presentation [")


# def test_repr_slides(slides: Slides):
#     assert repr(slides) == "<Slides>"
