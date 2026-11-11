# SPDX-License-Identifier: Apache-2.0
"""D1: bind upstream routing to the plugin's load-time Hermes profile."""
from __future__ import annotations

import logging
import uuid
from pathlib import Path

import pytest
from hermes_constants import reset_hermes_home_override, set_hermes_home_override

from test_desk_bridge_security import (
    API_KEY, Upstream, connect, make_runtime, new_session, request, running, session_headers,
)
from desk_plugin.desk_bridge import DeskError, check_api_server_url


@pytest.mark.parametrize("profile", ["default", "sidecar-test"])
@pytest.mark.parametrize("explicit", [False, True])
def test_d1_routes_own_profile_after_scope_ends(tmp_path, monkeypatch, caplog, profile, explicit):
    # The process serves default; plugin load is context-scoped to a named profile.
    root = tmp_path / "root"
    monkeypatch.setenv("HERMES_HOME", str(root))
    home = root if profile == "default" else root / "profiles" / profile
    upstream = Upstream(status=401)
    caplog.set_level(logging.DEBUG)
    prefix = "" if profile == "default" else f"/p/{profile}"
    token = set_hermes_home_override(home)
    try:
        rt = make_runtime(tmp_path / "data", api_url=upstream.url + (prefix if explicit else ""))
    finally:
        reset_hermes_home_override(token)
    try:
        with running(rt):
            grant = connect(rt)["grant"]
            pane = str(uuid.uuid4())
            _, _, session = new_session(rt, grant, pane)
            status, headers, body = request(
                rt, "POST", "/v1/chat/completions", {"messages": []},
                headers=session_headers(session, pane))
            assert status == 401
            seen = upstream.requests[-1]
            assert seen["path"] == prefix + "/v1/chat/completions"
            assert seen["headers"]["Authorization"] == f"Bearer {API_KEY}"
            assert API_KEY not in str(headers) and API_KEY.encode() not in body
            assert API_KEY not in caplog.text
    finally:
        upstream.close()


@pytest.mark.parametrize("path", [
    "/p/other", "/p/default", "/v1", "/p/sidecar-test/v1", "/p/sidecar-test/",
    "/p/%73idecar-test", "/p/sidecar-test/../other", "//p/sidecar-test",
])
def test_d1_rejects_foreign_and_arbitrary_paths(tmp_path, path):
    token = set_hermes_home_override(tmp_path / "profiles" / "sidecar-test")
    try:
        with pytest.raises(DeskError):
            check_api_server_url("http://127.0.0.1:8642" + path)
    finally:
        reset_hermes_home_override(token)


def test_d1_default_rejects_named_prefix(tmp_path, monkeypatch):
    monkeypatch.setenv("HERMES_HOME", str(tmp_path))
    with pytest.raises(DeskError):
        check_api_server_url("http://127.0.0.1:8642/p/other")
