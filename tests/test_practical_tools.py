"""Tests for practical tools (removed: merge_multiple_files, data_validation, conditional formatting — tools deleted in API simplification)."""
import os
import pytest
from openpyxl import Workbook

def _create_workbook(path, sheet_title, headers, rows):
    """Create a test workbook and save to path."""
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_title
    ws.append(headers)
    for row in rows:
        ws.append(row)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)
    return path
