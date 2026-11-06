from __future__ import annotations

import pytest

from xlsxpress import config
from xlsxpress.config import AppConfig, CleanupPolicy, DialogDefaults

pytestmark = pytest.mark.unit


class TestLoadConfig:
    def test_load_config_returns_defaults_without_a_config_file(self) -> None:
        assert config.load_config() == config.default_config()

    def test_load_config_returns_what_write_config_wrote(self) -> None:
        written = AppConfig(
            DialogDefaults(open_after=False, preview=True, nrows=5, to_temp=False, format_header=False),
            CleanupPolicy(max_age_days=3, max_files=7),
        )
        config.write_config(written)

        assert config.load_config() == written

    def test_load_config_ignores_unknown_keys_and_values_of_the_wrong_type(self) -> None:
        path = config.get_config_path()
        path.parent.mkdir(parents=True)
        path.write_text('[dialog]\nnrows = "many"\ncolor = "red"\n[cleanup]\nmax_files = 3\n', encoding="utf-8")

        loaded = config.load_config()

        assert loaded.dialog == DialogDefaults()
        assert loaded.cleanup == CleanupPolicy(max_files=3)

    def test_load_config_returns_defaults_for_an_invalid_file(self) -> None:
        path = config.get_config_path()
        path.parent.mkdir(parents=True)
        path.write_text("[dialog\n", encoding="utf-8")

        assert config.load_config() == config.default_config()
