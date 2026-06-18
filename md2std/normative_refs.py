# -*- coding: utf-8 -*-
"""Normative-reference registration and alias helpers."""

from __future__ import annotations

from dataclasses import dataclass
import re


@dataclass(frozen=True)
class NormativeRefEntry:
    target: str
    content: str
    code: str
    name: str
    explicit: bool
    marker_end: int = 0


_REGISTRATION_RE = re.compile(r"^\s*\{\{\s*std\s*:\s*([^{}]+?)\s*\}\}\s*(.*)$")
_ENTRY_RE = re.compile(
    r"^\s*([A-Z][A-Z0-9/]*(?:[ \t]+[A-Z0-9][A-Z0-9/.\-—–:]*)*)"
    r"(?:[ \t]{2,}|　+)(.+)$"
)
_SPACE_RE = re.compile(r"[ \t　]+")
_DOMESTIC_DATED_RE = re.compile(
    r"^((?:GB(?:/[TZ])?|GJB|T/[A-Z0-9]+|DB\d{2}/T|[A-Z]{2,5}/T)\s+.+?)([—\-–])(\d{4})$"
)
_FOREIGN_COLON_DATED_RE = re.compile(r"^(.+):(\d{4})$")
_DOMESTIC_BAD_YEAR_RE = re.compile(
    r"\b((?:GB(?:/[TZ])?|GJB|T/[A-Z0-9]+|DB\d{2}/T|[A-Z]{2,5}/T)\s+"
    r"[A-Z0-9][A-Z0-9/.\-—–:]*?)([\-–])(\d{4})\b"
)


def normalize_standard_id(value: str) -> str:
    return _SPACE_RE.sub(" ", (value or "").strip())


def parse_ref_registration(text: str) -> NormativeRefEntry | None:
    match = _REGISTRATION_RE.match(text or "")
    if not match:
        return None
    target = normalize_standard_id(match.group(1))
    content = match.group(2).strip()
    code, name = split_normative_ref_content(content)
    return NormativeRefEntry(
        target=target,
        content=content,
        code=code or target,
        name=name,
        explicit=True,
        marker_end=match.end(0) - len(match.group(2)),
    )


def parse_implicit_ref_entry(text: str) -> NormativeRefEntry | None:
    code, name = split_normative_ref_content(text)
    if not code:
        return None
    return NormativeRefEntry(
        target=code,
        content=(text or "").strip(),
        code=code,
        name=name,
        explicit=False,
    )


def split_normative_ref_content(text: str) -> tuple[str, str]:
    match = _ENTRY_RE.match(text or "")
    if not match:
        return "", ""
    return normalize_standard_id(match.group(1)), match.group(2).strip()


def standard_aliases(*values: str) -> set[str]:
    aliases: set[str] = set()
    for value in values:
        code = normalize_standard_id(value)
        if not code:
            continue
        aliases.add(code)
        domestic = _DOMESTIC_DATED_RE.match(code)
        if domestic:
            base, sep, year = domestic.groups()
            base = normalize_standard_id(base)
            aliases.add(base)
            aliases.add(f"{base}—{year}")
            if sep != "—":
                aliases.add(f"{base}{sep}{year}")
            continue
        foreign = _FOREIGN_COLON_DATED_RE.match(code)
        if foreign:
            aliases.add(normalize_standard_id(foreign.group(1)))
    return aliases


def domestic_year_separator_issues(text: str) -> list[tuple[str, str]]:
    issues: list[tuple[str, str]] = []
    for match in _DOMESTIC_BAD_YEAR_RE.finditer(text or ""):
        issues.append((normalize_standard_id("%s%s%s" % match.groups()), match.group(2)))
    return issues
