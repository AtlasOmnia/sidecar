# SPDX-License-Identifier: Apache-2.0
"""Load this repo's Hermes plugin directory as an importable package for tests.

The plugin lives in plugins/atlas-sidecar-<app>/ (a directory name Python cannot import directly).
Tests import it as ``desk_plugin`` (and its companion as ``desk_plugin.companion``) without Hermes.
"""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
_candidates = sorted((REPO / "plugins").glob("atlas-sidecar-*"))
assert len(_candidates) == 1, _candidates
PLUGIN_DIR = _candidates[0]
PACKAGE = "desk_plugin"


def load(name: str = PACKAGE):
    """Import (once) and return the plugin package under ``name``."""
    if name in sys.modules:
        return sys.modules[name]
    spec = importlib.util.spec_from_file_location(
        name, PLUGIN_DIR / "__init__.py", submodule_search_locations=[str(PLUGIN_DIR)])
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    try:
        spec.loader.exec_module(module)
    except BaseException:
        sys.modules.pop(name, None)
        raise
    return module


def fresh(name: str):
    """A second, independent copy of the plugin (a new module namespace, like a second profile)."""
    for key in [k for k in sys.modules if k == name or k.startswith(name + ".")]:
        del sys.modules[key]
    return load(name)
