"""Parsing logic for extracting product rows from text."""

from __future__ import annotations

import re
from typing import Dict, List, Pattern


def parse_items(text: str, template: Pattern[str] | None = None) -> List[Dict[str, str]]:
    """Parse OCR text and extract product items.

    Parameters
    ----------
    text: str
        Cleaned OCR text.
    template: Pattern[str], optional
        Regex pattern describing a single line. Defaults to
        ``<sku> <name> <qty>`` structure.
    """
    if template is None:
        template = re.compile(r"(\S+)\s+([\w\s]+?)\s+(\d+)")

    items: List[Dict[str, str]] = []
    for line in text.splitlines():
        match = template.search(line)
        if match:
            sku, name, qty = match.groups()
            items.append({"sku": sku, "name": name.strip(), "quantity": qty})
    return items
