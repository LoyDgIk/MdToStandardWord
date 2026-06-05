# -*- coding: utf-8 -*-
"""Package resource helpers for bundled md2std assets."""

from __future__ import annotations

from contextlib import ExitStack
from importlib import resources
from pathlib import Path

_RESOURCE_STACK = ExitStack()


def template_path(filename: str) -> str:
    """Return a filesystem path for a bundled template asset."""
    candidate = resources.files("md2std").joinpath("templates", filename)
    if not candidate.is_file():
        raise FileNotFoundError("找不到模板资源：%s" % filename)
    return str(_RESOURCE_STACK.enter_context(resources.as_file(candidate)))


def template_candidates(*filenames: str) -> list[str]:
    paths = []
    for filename in filenames:
        try:
            paths.append(template_path(filename))
        except FileNotFoundError:
            continue
    return paths


def package_root() -> Path:
    return Path(__file__).resolve().parent
