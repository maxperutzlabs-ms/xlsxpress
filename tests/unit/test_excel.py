from __future__ import annotations

import sys
from pathlib import Path

import pytest

from xlsxpress.core import excel

pytestmark = pytest.mark.unit


class TestShowInFolder:
    @pytest.mark.parametrize(
        ("platform", "expected"),
        [
            ("win32", ["explorer", "/select,", str(Path("folder") / "data.xlsx")]),
            ("darwin", ["open", "-R", str(Path("folder") / "data.xlsx")]),
            ("linux", ["xdg-open", "folder"]),
        ],
    )
    def test_show_in_folder_starts_the_platform_file_manager(
        self, monkeypatch: pytest.MonkeyPatch, platform: str, expected: list[str]
    ) -> None:
        commands: list[list[str]] = []
        monkeypatch.setattr(sys, "platform", platform)
        monkeypatch.setattr(excel.subprocess, "Popen", commands.append)

        excel.show_in_folder(Path("folder") / "data.xlsx")

        assert commands == [expected]
