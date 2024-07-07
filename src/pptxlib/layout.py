# """
# CustomLayoutに関連するモジュール
# """


# def copy_layout(slide, name=None, replace=True):
#     """指定するスライドのCustomLayoutをコピーして返す．

#     Parameters
#     ----------
#     slide : xlviews.powerpoint.main.Slide
#         スライドオブジェクト
#     name : str, optional
#         CustomLayoutの名前
#     replace : bool, optional
#         スライドのCustomLayoutをコピーしたものに
#         置き換えるか

#     Returns
#     -------
#     layout
#     """
#     layouts = slide.parent.api.SlideMaster.CustomLayouts
#     slide.api.CustomLayout.Copy()
#     layout = layouts.Paste()
#     if name:
#         layout.Name = name
#     if replace:
#         slide.api.CustomLayout = layout
#     return layout
