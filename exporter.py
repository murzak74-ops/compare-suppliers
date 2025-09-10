"""Functions for exporting parsed data to Excel."""

from __future__ import annotations

from typing import Dict, List

from openpyxl import Workbook


def export_to_excel(items: List[Dict[str, str]], output_path: str) -> None:
    """Save parsed items into an Excel file using openpyxl."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Артикул", "Наименование", "Количество"])
    for item in items:
        ws.append([item.get("sku", ""), item.get("name", ""), item.get("quantity", "")])
    wb.save(output_path)
