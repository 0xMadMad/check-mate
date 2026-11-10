"""Regression tests for CheckMate exports (run with pytest from the repo root)."""

import os
import sys
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import core  # noqa: E402
import models  # noqa: E402


def _manager_with_empty_and_full_lists():
    mgr = models.ToDoManager()
    mgr.add_tasklist(models.TaskList("Empty"))
    work = models.TaskList("Work")
    work.add_task(models.Task("T1", "desc", 5))
    mgr.add_tasklist(work)
    return mgr


def test_export_to_excel_empty_tasklist_has_no_extra_column(tmp_path):
    from openpyxl import load_workbook

    mgr = _manager_with_empty_and_full_lists()
    with (
        patch("builtins.input", side_effect=["Y", "checkmate_test.xlsx"]),
        patch.object(core, "DEFAULT_PATH", str(tmp_path) + os.sep),
    ):
        core.export_to_excel(mgr)
    ws = load_workbook(tmp_path / "checkmate_test.xlsx").active
    rows = list(ws.iter_rows(values_only=True))
    assert [c for c in rows[0] if c is not None] == [
        "TaskList Name",
        "Task Name",
        "Task Description",
        "Task Importance",
        "Task Status",
    ]
    assert ws.max_column == 5
    empty_row = rows[1]
    assert empty_row[0] == "Empty"
    assert all(c is None for c in empty_row[1:5])
