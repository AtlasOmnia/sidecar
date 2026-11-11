# SPDX-License-Identifier: Apache-2.0
"""Make ``desk_plugin`` (this repo's plugin directory) importable for every test module, and keep every
test away from the real Hermes home (plugin data goes to a temporary HERMES_HOME)."""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))

import desk_loader  # noqa: E402

desk_loader.load()


@pytest.fixture(autouse=True)
def _isolated_hermes_home(tmp_path_factory, monkeypatch):
    home = tmp_path_factory.mktemp("hermes-home")
    monkeypatch.setenv("HERMES_HOME", str(home))
    monkeypatch.delenv("_HERMES_GATEWAY", raising=False)
    yield home
