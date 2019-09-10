import re

p = re.compile(r"(\?)?\.")


def get_obj_attr(obj, field):
    index = 0
    for match in p.finditer(field):
        start, end, s = match.start(), match.end(), match.group()
        obj = getattr(obj, field[index:start])
        if s == "?." and obj is None:
            return obj
        index = end
    obj = getattr(obj, field[index:])
    return obj
