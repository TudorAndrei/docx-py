"""Utility functions module."""

import base64
import re
from collections.abc import Callable
from typing import Any, TypeVar
from urllib.parse import urljoin

T = TypeVar("T")


def escape_class_name(class_name: str | None) -> str | None:
    """Escape class name for CSS."""
    if not class_name:
        return None
    return re.sub(r"[ .]+", "-", class_name).replace("&", "and").lower()


def enclose_font_family(font_family: str) -> str:
    """Enclose font family in quotes if needed."""
    if re.match(r'^[^"\'].*\s.*[^"\']$', font_family):
        return f"'{font_family}'"
    return font_family


def split_path(path: str) -> tuple[str, str]:
    """Split path into folder and file name."""
    si = path.rfind("/") + 1
    folder = "" if si == 0 else path[:si]
    file_name = path if si == 0 else path[si:]
    return (folder, file_name)


def resolve_path(path: str, base: str) -> str:
    """Resolve relative path against base path."""
    try:
        prefix = "http://docx/"
        url = urljoin(prefix + base, path)
        return url[len(prefix) :]
    except (ValueError, AttributeError):
        return f"{base}{path}"


def key_by(array: list[T], by: Callable[[T], Any]) -> dict[Any, T]:
    """Create dictionary keyed by function result."""
    result = {}
    for x in array:
        key = by(x)
        result[key] = x
    return result


def is_object(item: Any) -> bool:  # noqa: ANN401
    """Check if item is a dictionary."""
    return item is not None and isinstance(item, dict)


def is_string(item: Any) -> bool:  # noqa: ANN401
    """Check if item is a string."""
    return isinstance(item, str)


def merge_deep(target: dict, *sources: dict) -> dict:
    """Deep merge dictionaries."""
    if not sources:
        return target

    source = sources[0]

    if is_object(target) and is_object(source):
        for key in source:
            if is_object(source[key]):
                val = target.get(key, {})
                if not is_object(val):
                    val = {}
                target[key] = val
                merge_deep(val, source[key])
            else:
                target[key] = source[key]

    return merge_deep(target, *sources[1:])


CSS_RULE_PARTS_COUNT = 2


def parse_css_rules(text: str) -> dict[str, str]:
    """Parse CSS rules from text."""
    result = {}
    for rule in text.split(";"):
        parts = rule.split(":", 1)
        if len(parts) == CSS_RULE_PARTS_COUNT:
            key = parts[0].strip()
            val = parts[1].strip()
            result[key] = val
    return result


def format_css_rules(style: dict[str, str]) -> str:
    """Format CSS rules as string."""
    return ";".join(f"{k}: {v}" for k, v in style.items())


def as_array(val: Any) -> list:  # noqa: ANN401
    """Convert value to array."""
    return val if isinstance(val, list) else [val]


def clamp(val: float, min_val: float, max_val: float) -> float:
    """Clamp value between min and max."""
    return min(max(min_val, val), max_val)


def bytes_to_base64(data: bytes) -> str:
    """Convert bytes to base64 encoded string."""
    return base64.b64encode(data).decode("utf-8")

