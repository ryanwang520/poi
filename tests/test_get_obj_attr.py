from poi.utils import get_obj_attr
from types import SimpleNamespace as Obj


def test_get_obj_attr():
    obj = Obj(foo=Obj(bar=Obj(moo=14)))
    assert get_obj_attr(obj, "foo.bar.moo") == 14

    obj = Obj(foo=None)
    assert get_obj_attr(obj, "foo?.bar") is None

    obj = Obj(foo=Obj(bar=None))
    assert get_obj_attr(obj, "foo?.bar?.moo") is None
