# SPDX-License-Identifier: Apache-2.0
# Copyright 2026 Atlas Omnia
"""App constants for this plugin's desk bridge (desk_bridge.py follows the same design across the three plugins)."""
from __future__ import annotations

from .companion.handler import CompanionHandler
from .companion.server import WordCompanionServer as CompanionServer

APP = "word"
APP_TITLE = "Word"
DOCUMENT_NOUN = "document"
PLUGIN_ID = "atlas-sidecar-word"
COMMAND = "sidecar-word"
SERVICE = "hermes-desk-word"
TOOL_PREFIX = "sidecar_word_"
TOOLSET = "sidecar-word"
# Own range: never Workbench's 47615-47619, never another Sidecar app's range.
PORTS = tuple(range(47630, 47635))
# The plugin serves the bundled sidebar from its own HTTPS loopback bridge.
DEFAULT_ADDIN_BASE = f"https://localhost:{PORTS[0]}/{COMMAND}/"

__all__ = ["APP", "APP_TITLE", "DOCUMENT_NOUN", "PLUGIN_ID", "COMMAND", "SERVICE", "TOOL_PREFIX", "TOOLSET",
           "PORTS", "DEFAULT_ADDIN_BASE", "CompanionHandler", "CompanionServer"]
