from types import SimpleNamespace as Obj

from poi.utils import get_obj_attr


def test_get_obj_attr():
    obj = Obj(foo=Obj(bar=Obj(moo=14)))
    assert get_obj_attr(obj, "foo.bar.moo") == 14

    obj = Obj(foo=None)
    assert get_obj_attr(obj, "foo?.bar") is None

    obj = Obj(foo=Obj(bar=None))
    assert get_obj_attr(obj, "foo?.bar?.moo") is None


def test_get_obj_attr_for_dict():
    obj = {"foo": {"bar": {"moo": 14}}}
    assert get_obj_attr(obj, "foo.bar.moo") == 14

    obj = {"foo": None}
    assert get_obj_attr(obj, "foo?.bar") is None

    obj = {"foo": {"bar": None}}
    assert get_obj_attr(obj, "foo?.bar?.moo") is None
