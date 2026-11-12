# SPDX-License-Identifier: Apache-2.0
# Copyright 2026 Atlas Omnia
"""App constants for this plugin's desk bridge (desk_bridge.py is shared verbatim by the three plugins)."""
from __future__ import annotations

from .companion.handler import CompanionHandler
from .companion.server import ExcelCompanionServer as CompanionServer

APP = "excel"
APP_TITLE = "Excel"
DOCUMENT_NOUN = "workbook"
PLUGIN_ID = "atlas-sidecar-excel"
COMMAND = "sidecar-excel"
SERVICE = "hermes-desk-excel"
TOOL_PREFIX = "sidecar_excel_"
TOOLSET = "sidecar-excel"
# Own range: never Workbench's 47615-47619, never another Sidecar app's range.
PORTS = tuple(range(47640, 47645))
# Default sidebar base: the plugin serves the bundled page from its own loopback bridge.
# The runtime replaces the default port with the first port it successfully binds.
DEFAULT_ADDIN_BASE = f"https://localhost:{PORTS[0]}/{COMMAND}/"

__all__ = ["APP", "APP_TITLE", "DOCUMENT_NOUN", "PLUGIN_ID", "COMMAND", "SERVICE", "TOOL_PREFIX", "TOOLSET",
           "PORTS", "DEFAULT_ADDIN_BASE", "CompanionHandler", "CompanionServer"]
