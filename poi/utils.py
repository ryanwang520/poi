import re
from typing import Any

p = re.compile(r"(\?)?\.")


def get_obj_attr(obj: object, field: str) -> Any:
    """
    Retrieve nested attribute or dictionary key from an object based on a field string.

    Parameters:
    - obj (object): The object or dictionary from which to extract the attribute or key.
    - field (str): The string specifying the nested attributes or keys, separated by '.'
                   and/or '?.' where '?.' performs safe navigation, returning None if the
                   attribute is not found.

    Returns:
    - Any: The value of the nested attribute or key. Returns None if the attribute is not
           found and safe navigation '?.' is used.

    Examples:
    >>> obj = {'a': {'b': {'c': 1}}}
    >>> get_obj_attr(obj, 'a.b.c')
    1

    >>> class MyClass:
    ...     def __init__(self):
    ...         self.a = {'b': 2}
    >>> obj = MyClass()
    >>> get_obj_attr(obj, 'a.b')
    2

    >>> get_obj_attr(obj, 'a?.c')
    None

    """
    index = 0
    for match in p.finditer(field):
        start, end, s = match.start(), match.end(), match.group()
        if isinstance(obj, dict):
            obj = obj.get(field[index:start])
        else:
            obj = getattr(obj, field[index:start])

        if s == "?." and obj is None:
            return obj
        index = end
    if isinstance(obj, dict):
        obj = obj[field[index:]]
    else:
        obj = getattr(obj, field[index:])
    return obj
