# SPDX-License-Identifier: Apache-2.0
# Copyright 2026 Atlas Omnia
"""App constants for this plugin's desk bridge (desk_bridge.py follows the same design across the three plugins)."""
from __future__ import annotations

from .companion.handler import CompanionHandler
from .companion.server import PowerPointCompanionServer as CompanionServer

APP = "powerpoint"
APP_TITLE = "PowerPoint"
DOCUMENT_NOUN = "presentation"
PLUGIN_ID = "atlas-sidecar-powerpoint"
COMMAND = "sidecar-powerpoint"
SERVICE = "hermes-desk-powerpoint"
TOOL_PREFIX = "sidecar_powerpoint_"
TOOLSET = "sidecar-powerpoint"
# Own range: never Workbench's 47615-47619, never another Sidecar app's range.
PORTS = tuple(range(47650, 47655))
# Public sidebar origin; first exercised at publication. Override addin_base for your own site.
DEFAULT_ADDIN_BASE = "https://atlasomnia.github.io/sidecar-powerpoint/"

__all__ = ["APP", "APP_TITLE", "DOCUMENT_NOUN", "PLUGIN_ID", "COMMAND", "SERVICE", "TOOL_PREFIX", "TOOLSET",
           "PORTS", "DEFAULT_ADDIN_BASE", "CompanionHandler", "CompanionServer"]
