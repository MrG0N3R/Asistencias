from __future__ import annotations

import re

def clean(value) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def normalize_employee_id(value) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return clean(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = clean(value)
    numeric_text = re.fullmatch(r"(\d+)\.0+", text)
    return numeric_text.group(1) if numeric_text else text
