from functools import wraps
import xlwings as xw


def wait_updating(func):
    """
    処理が終了するまで，画面の更新をやめる．
    """
    @wraps(func)
    def _func(*args, **kwargs):
        active = xw.apps.active
        if active:  # for powerpoint
            is_updating = active.screen_updating
            active.screen_updating = False
        else:
            is_updating = None
        try:
            result = func(*args, **kwargs)
        finally:
            if active:
                active.screen_updating = is_updating
        return result
    return _func


def api(func):
    """
    第一引数がxlwingsオブジェクトの場合にwin32comオブジェクトにキャストする．
    """
    @wraps(func)
    def _func(obj, *args, **kwargs):
        if hasattr(obj, 'api'):
            obj = obj.api
        return func(obj, *args, **kwargs)
    return _func
