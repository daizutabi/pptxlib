from importlib.metadata import version


def test_pptxlib():
    assert version("pptxlib").count(".") == 2
