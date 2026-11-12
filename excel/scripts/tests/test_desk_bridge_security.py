# SPDX-License-Identifier: Apache-2.0
"""Security behaviour of the desk bridge that runs inside this Hermes plugin (RED/GREEN suite).

App-agnostic: every app constant comes from the plugin's own desk_app module, so this file is the same
in the Word, Excel and PowerPoint repositories.
"""
from __future__ import annotations

import http.client
import io
import json
import logging
import os
import socket
import ssl
import threading
import time
import uuid
from contextlib import contextmanager, redirect_stderr
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import pytest

import desk_loader

PKG = desk_loader.load()
from desk_plugin import desk_app  # noqa: E402
from desk_plugin.desk_bridge import (  # noqa: E402
    GRANT_HEADER, ApprovalStore, DeskError, DeskRuntime, KeyProvider, _render_manifest, check_api_server_url, origin_of)

APP = desk_app.APP
ADDIN_BASE = "https://sidebar.example.test/office/originals/%s/" % APP
ORIGIN = origin_of(ADDIN_BASE)
API_KEY = "sk-desk-test-" + "Z9" * 24  # a recognisable fake API_SERVER_KEY

# Workbench (hermes-office-desktop @ b2d6547) plugin.yaml provides_tools, pinned verbatim.
WORKBENCH_TOOLS = frozenset("""
excel_get_workbook_overview excel_read_range excel_summarize_range_stats excel_write_range excel_format_range
excel_conditional_format_range excel_freeze_panes excel_set_data_validation excel_create_sheet excel_delete_sheet
excel_activate_sheet excel_insert_table excel_sort_table excel_filter_table excel_create_chart excel_find_replace
excel_clear_range excel_autofill_formula excel_rename_sheet excel_copy_sheet excel_protect_sheet
excel_unprotect_sheet excel_create_pivot_table excel_run_calculation excel_add_comment excel_clear_comments
excel_import_csv_to_sheet word_get_document_overview word_read_selection word_read_document_text
word_apply_edit_plan word_insert_text word_replace_selection word_find_replace word_format_selection
word_get_document_structure word_apply_style word_insert_ooxml word_insert_comment
powerpoint_get_presentation_overview powerpoint_read_selected_slides powerpoint_read_all_slides
powerpoint_create_slide powerpoint_create_text_slide powerpoint_add_text_box powerpoint_update_slide_text
powerpoint_add_speaker_notes powerpoint_add_shape powerpoint_read_slide_content powerpoint_delete_slide
powerpoint_duplicate_slide powerpoint_reorder_slide powerpoint_find_and_replace_text
powerpoint_list_slide_layouts powerpoint_set_slide_layout powerpoint_format_text powerpoint_add_image
onenote_get_context onenote_list_sections onenote_list_pages onenote_get_page onenote_create_page onenote_append
onenote_insert_table onenote_open_page brain_sync brain_ask brain_links brain_link_both_ways brain_decision_page
brain_make_deliverables brain_why brain_resurface brain_from_email outlook_get_item outlook_insert
outlook_draft_reply outlook_draft_new calendar_get_event calendar_update_draft calendar_draft_new office_connect
office_status office_disconnect office_undo office_insert_image
""".split())
WORKBENCH_PREFIXES = ("excel_", "word_", "powerpoint_", "onenote_", "outlook_", "office_", "brain_", "calendar_")
WORKBENCH_PORTS = set(range(47615, 47620))


# ------------------------------------------------------------------------------------------------
# helpers
# ------------------------------------------------------------------------------------------------

def _free_ports(n: int) -> tuple:
    socks, ports = [], []
    for _ in range(n):
        s = socket.socket()
        s.bind(("127.0.0.1", 0))
        socks.append(s)
        ports.append(s.getsockname()[1])
    for s in socks:
        s.close()
    return tuple(ports)


class _AppCfg:
    """desk_app with free test ports (the real range is asserted separately)."""

    def __init__(self, ports: tuple) -> None:
        for name in dir(desk_app):
            if name.isupper() or name in ("CompanionServer", "CompanionHandler"):
                setattr(self, name, getattr(desk_app, name))
        self.PORTS = ports


class Upstream:
    """A fake Hermes API server: records requests, answers with a short SSE stream."""

    def __init__(self, status: int = 200) -> None:
        self.requests: list[dict] = []
        outer = self

        class H(BaseHTTPRequestHandler):
            protocol_version = "HTTP/1.1"

            def log_message(self, *a):  # noqa: D401
                return

            def do_POST(self):  # noqa: N802
                body = self.rfile.read(int(self.headers.get("Content-Length", "0")))
                outer.requests.append({"path": self.path, "headers": dict(self.headers), "body": body})
                payload = b"data: {\"choices\":[{\"delta\":{\"content\":\"hi\"}}]}\n\ndata: [DONE]\n\n"
                self.send_response(status)
                self.send_header("Content-Type", "text/event-stream")
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)

        self.httpd = ThreadingHTTPServer(("127.0.0.1", 0), H)
        self.url = f"http://127.0.0.1:{self.httpd.server_address[1]}"
        threading.Thread(target=self.httpd.serve_forever, daemon=True).start()

    def close(self) -> None:
        self.httpd.shutdown()
        self.httpd.server_close()


def make_runtime(data_dir: Path, *, api_url: str = "http://127.0.0.1:9", choice: str | None = None,
                 key: str | None = API_KEY, ports: tuple | None = None, gateway: bool = False) -> DeskRuntime:
    return DeskRuntime(appcfg=_AppCfg(ports or _free_ports(3)), data_dir=data_dir, addin_base=ADDIN_BASE,
                       api_server_url=api_url, key_provider=KeyProvider(lambda: key),
                       ask_choice=(lambda q, c: choice), in_gateway=lambda: gateway)


def request(runtime: DeskRuntime, method: str, path: str, body=None, headers=None, origin: str | None = ORIGIN,
            host: str | None = None):
    port = runtime.bridge.port
    conn = http.client.HTTPSConnection("localhost", port, timeout=10,
                                      context=ssl.create_default_context(cafile=str(runtime.store.data_dir / "tls/ca.pem")))
    hdrs = {"Host": host or f"localhost:{port}"}
    if origin is not None:
        hdrs["Origin"] = origin
    data = None
    if body is not None:
        data = json.dumps(body).encode()
        hdrs["Content-Type"] = "application/json"
        hdrs["Content-Length"] = str(len(data))
    hdrs.update(headers or {})
    conn.request(method, path, body=data, headers=hdrs)
    resp = conn.getresponse()
    raw = resp.read()
    conn.close()
    try:
        parsed = json.loads(raw.decode()) if raw else None
    except ValueError:
        parsed = raw
    return resp.status, dict(resp.getheaders()), parsed


def connect(runtime: DeskRuntime, decision: str | None = "allow") -> dict:
    status, _, created = request(runtime, "POST", f"/{APP}/connect", {"client": {"host": APP, "platform": "Mac"}})
    assert status == 201, created
    if decision == "allow":
        runtime.store.decide(True, created["code"])
    elif decision == "deny":
        runtime.store.decide(False, created["code"])
    status, _, polled = request(runtime, "POST", f"/{APP}/connect/status",
                                {"requestId": created["requestId"], "pollSecret": created["pollSecret"]})
    assert status == 200, polled
    return {**created, **polled}


def new_session(runtime: DeskRuntime, grant: str | None, pane_id: str | None = None):
    headers = {GRANT_HEADER: grant} if grant else {}
    return request(runtime, "POST", f"/{APP}/session",
                   {"protocolVersion": "1.0", "documentId": str(uuid.uuid4()), "paneId": pane_id or str(uuid.uuid4())},
                   headers=headers)


def session_headers(created: dict, pane_id: str) -> dict:
    return {"X-Hermes-Session-Id": created["sessionId"], "X-Hermes-Session-Key": created["sessionKey"],
            "X-Hermes-Pane-Id": pane_id}


@contextmanager
def running(runtime: DeskRuntime):
    runtime.ensure_started()
    try:
        yield runtime
    finally:
        runtime.stop()


# ------------------------------------------------------------------------------------------------
# bind / ports / names
# ------------------------------------------------------------------------------------------------

def test_port_range_is_loopback_and_never_overlaps_workbench() -> None:
    assert len(desk_app.PORTS) == 5
    assert not set(desk_app.PORTS) & WORKBENCH_PORTS
    ranges = {"word": range(47630, 47635), "excel": range(47640, 47645), "powerpoint": range(47650, 47655)}
    assert tuple(desk_app.PORTS) == tuple(ranges[APP])


def test_bridge_binds_127_0_0_1_only(tmp_path: Path) -> None:
    with running(make_runtime(tmp_path)) as rt:
        server = rt.bridge._server
        assert server.server_address()[0] == "127.0.0.1"
        assert server._server.socket.getsockname()[0] == "127.0.0.1"
        # Not reachable on another local address of this machine.
        try:
            other = socket.gethostbyname(socket.gethostname())
        except OSError:
            other = None
        if other and not other.startswith("127."):
            with pytest.raises(OSError):
                socket.create_connection((other, rt.bridge.port), timeout=1).close()


def test_bridge_uses_first_free_port_of_its_range(tmp_path: Path) -> None:
    ports = _free_ports(2)
    blocker = socket.socket()
    blocker.bind(("127.0.0.1", ports[0]))
    blocker.listen(1)
    try:
        with running(make_runtime(tmp_path, ports=ports)) as rt:
            assert rt.bridge.port == ports[1]
    finally:
        blocker.close()


def test_api_server_url_must_be_loopback() -> None:
    assert check_api_server_url("http://127.0.0.1:8642") == ("127.0.0.1", 8642)
    for bad in ("http://10.0.0.2:8642", "https://127.0.0.1:8642", "http://example.com:8642",
                "http://127.0.0.1:8642/v1", "http://user:pw@127.0.0.1:8642", "http://127.0.0.1"):
        with pytest.raises(DeskError):
            check_api_server_url(bad)


def test_tool_names_are_prefixed_and_disjoint_from_workbench() -> None:
    schemas = json.loads((desk_loader.PLUGIN_DIR / "tool_schemas.json").read_text(encoding="utf-8"))
    names = {s["name"] for s in schemas}
    assert names and all(n.startswith(desk_app.TOOL_PREFIX) for n in names)
    assert desk_app.TOOL_PREFIX == f"sidecar_{APP}_"
    assert not names & WORKBENCH_TOOLS
    assert not any(n.startswith(WORKBENCH_PREFIXES) for n in names)
    manifest = (desk_loader.PLUGIN_DIR / "plugin.yaml").read_text(encoding="utf-8")
    declared = {line.strip()[2:] for line in manifest.splitlines() if line.strip().startswith("- sidecar_")}
    assert declared == names


# ------------------------------------------------------------------------------------------------
# origin and host
# ------------------------------------------------------------------------------------------------

@pytest.mark.parametrize("origin", [None, "", "https://evil.example", "https://localhost:3001", "null",
                                    ORIGIN + ".evil.example", ORIGIN.replace("https", "http")])
def test_wrong_or_missing_origin_is_rejected_everywhere(tmp_path: Path, origin) -> None:
    with running(make_runtime(tmp_path)) as rt:
        for method, path, body in [("GET", f"/{APP}/hello", None),
                                   ("POST", f"/{APP}/connect", {"client": {}}),
                                   ("POST", f"/{APP}/connect/status", {"requestId": "x", "pollSecret": "y"}),
                                   ("POST", f"/{APP}/session", {"protocolVersion": "1.0", "documentId": "d",
                                                               "paneId": str(uuid.uuid4())}),
                                   ("POST", "/v1/chat/completions", {"messages": []})]:
            status, headers, _ = request(rt, method, path, body, origin=origin)
            assert status == 403, (method, path, status)
            assert "Access-Control-Allow-Origin" not in headers
        status, headers, _ = request(rt, "OPTIONS", f"/{APP}/connect", origin=origin,
                                     headers={"Access-Control-Request-Method": "POST"})
        assert status == 403 and "Access-Control-Allow-Origin" not in headers
        assert rt.store.pending() == []


def test_allowed_origin_gets_exact_cors(tmp_path: Path) -> None:
    with running(make_runtime(tmp_path)) as rt:
        status, headers, body = request(rt, "GET", f"/{APP}/hello")
        assert status == 200 and headers["Access-Control-Allow-Origin"] == ORIGIN
        assert body["service"] == desk_app.SERVICE and body["app"] == APP
        status, headers, _ = request(rt, "OPTIONS", f"/{APP}/session",
                                     headers={"Access-Control-Request-Method": "POST",
                                              "Access-Control-Request-Private-Network": "true"})
        assert status == 204 and headers["Access-Control-Allow-Origin"] == ORIGIN
        assert GRANT_HEADER in headers["Access-Control-Allow-Headers"]
        assert headers.get("Access-Control-Allow-Private-Network") == "true"


def test_unexpected_host_header_is_refused(tmp_path: Path) -> None:
    with running(make_runtime(tmp_path)) as rt:
        for host in ("127.0.0.1:%d" % rt.bridge.port, "evil.example", "127.0.0.1"):
            status, _, _ = request(rt, "GET", f"/{APP}/hello", host=host)
            assert status == 421
            status, _, _ = request(rt, "GET", f"/{desk_app.COMMAND}/taskpane.html", origin=None, host=host)
            assert status == 421


@pytest.mark.parametrize("configured_base", [desk_app.DEFAULT_ADDIN_BASE, ""])
def test_default_loopback_sidebar_is_served_from_the_bridge(tmp_path: Path, configured_base: str) -> None:
    rt = DeskRuntime(appcfg=_AppCfg(_free_ports(2)), data_dir=tmp_path,
                     addin_base=configured_base, api_server_url="http://127.0.0.1:9",
                     key_provider=KeyProvider(lambda: API_KEY), in_gateway=lambda: False)
    assert desk_app.DEFAULT_ADDIN_BASE == f"https://localhost:{desk_app.PORTS[0]}/{desk_app.COMMAND}/"
    with running(rt):
        assert rt.addin_base == f"https://localhost:{rt.bridge.port}/{desk_app.COMMAND}/"
        assert rt.allowed_origin == f"https://localhost:{rt.bridge.port}"
        status, headers, body = request(rt, "GET", f"/{desk_app.COMMAND}/taskpane.html", origin=None)
        assert status == 200
        assert isinstance(body, bytes)
        assert headers["Content-Type"].startswith("text/html")
        assert int(headers["Content-Length"]) == len(body)
        assert headers["Cache-Control"] == "no-cache"
        assert headers["X-Content-Type-Options"] == "nosniff"
        assert headers["Referrer-Policy"] == "same-origin"
        assert b"<html" in body.lower()
        policy = headers["Content-Security-Policy"]
        assert "script-src 'self' https://appsforoffice.microsoft.com" in policy
        assert "style-src 'self'" in policy
        assert "connect-src 'self'" in policy and "https://localhost:*" in policy
        assert "frame-ancestors" in policy and "https://*.office.com" in policy
        assert headers.get("X-Frame-Options", "").lower() != "deny"
        status, head_headers, head_body = request(rt, "HEAD", f"/{desk_app.COMMAND}/taskpane.html", origin=None)
        assert status == 200 and head_body is None
        assert int(head_headers["Content-Length"]) == len(body)


def test_sidebar_static_route_fails_closed(tmp_path: Path) -> None:
    rt = DeskRuntime(appcfg=_AppCfg(_free_ports(2)), data_dir=tmp_path,
                     addin_base=desk_app.DEFAULT_ADDIN_BASE, api_server_url="http://127.0.0.1:9",
                     key_provider=KeyProvider(lambda: API_KEY), in_gateway=lambda: False)
    with running(rt):
        prefix = f"/{desk_app.COMMAND}"
        for path in (f"{prefix}/missing.js", f"{prefix}/../desk_bridge.py",
                     f"{prefix}/%2e%2e/desk_bridge.py", f"{prefix}/file..txt",
                     f"{prefix}/assets%2ficon-32.png", f"{prefix}/%5cdesk_bridge.py",
                     f"{prefix}/%00", f"{prefix}//taskpane.html", f"{prefix}/%2ftaskpane.html"):
            status, _, _ = request(rt, "GET", path, origin=None)
            assert status == 404, path
        for method in ("POST", "PUT", "DELETE", "OPTIONS"):
            status, _, _ = request(rt, method, f"{prefix}/taskpane.html", origin=None)
            assert status == 405, method
        status, headers, _ = request(rt, "GET", f"{prefix}/taskpane.html", origin="https://evil.example")
        assert status == 403
        assert "Access-Control-Allow-Origin" not in headers


def test_sidebar_static_route_preserves_non_default_origin(tmp_path: Path) -> None:
    override = "https://sidebar.example.test/office/excel/"
    rt = DeskRuntime(appcfg=_AppCfg(_free_ports(2)), data_dir=tmp_path,
                     addin_base=override, api_server_url="http://127.0.0.1:9",
                     key_provider=KeyProvider(lambda: API_KEY), in_gateway=lambda: False)
    with running(rt):
        assert rt.addin_base == override
        assert rt.allowed_origin == origin_of(override)
        status, _, body = request(rt, "GET", f"/{desk_app.COMMAND}/taskpane.html", origin=None)
        assert status == 200 and isinstance(body, bytes) and b"<html" in body.lower()
        status, headers, _ = request(rt, "GET", f"/{desk_app.COMMAND}/taskpane.html", origin="https://evil.example")
        assert status == 403 and "Access-Control-Allow-Origin" not in headers


def test_manifest_command_writes_current_loopback_manifest(tmp_path: Path, capsys) -> None:
    rt = DeskRuntime(appcfg=_AppCfg(_free_ports(2)), data_dir=tmp_path,
                     addin_base=desk_app.DEFAULT_ADDIN_BASE, api_server_url="http://127.0.0.1:9",
                     key_provider=KeyProvider(lambda: API_KEY), in_gateway=lambda: False)
    with running(rt):
        slash_output = rt.slash("manifest")
        manifest_path = tmp_path / "manifest.xml"
        assert str(manifest_path) in slash_output
        assert 'mkdir -p "$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"' in slash_output
        assert '"$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef/atlas-sidecar-excel.xml"' in slash_output
        manifest = manifest_path.read_text(encoding="utf-8")
        assert f"https://localhost:{rt.bridge.port}/{desk_app.COMMAND}/taskpane.html" in manifest
        assert f"<AppDomain>https://localhost:{rt.bridge.port}</AppDomain>" in manifest
        assert "{{ADDIN_" not in manifest

        class Args:
            action = "manifest"
            code = None

        assert rt.cli_run(Args) == 0
        cli_output = capsys.readouterr().out
        assert str(manifest_path) in cli_output
        assert '"$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef/atlas-sidecar-excel.xml"' in cli_output


def test_manifest_fallback_renders_bundled_default_for_new_port(tmp_path: Path) -> None:
    installed_plugin = tmp_path / "installed-plugin"
    sidebar = installed_plugin / "sidebar"
    sidebar.mkdir(parents=True)
    template_path = desk_loader.PLUGIN_DIR.parents[1] / "manifest.template.xml"
    template = template_path.read_text(encoding="utf-8")
    default_manifest = (template.replace("{{ADDIN_BASE}}", desk_app.DEFAULT_ADDIN_BASE)
                        .replace("{{ADDIN_ORIGIN}}", origin_of(desk_app.DEFAULT_ADDIN_BASE)))
    (sidebar / "manifest.xml").write_text(default_manifest, encoding="utf-8")
    actual_base = f"https://localhost:47642/{desk_app.COMMAND}/"

    rendered = _render_manifest(installed_plugin, actual_base, desk_app.DEFAULT_ADDIN_BASE)

    assert actual_base + "taskpane.html" in rendered
    assert "https://localhost:47640/sidecar-excel/" not in rendered
    assert "<AppDomain>https://localhost:47642</AppDomain>" in rendered
    assert "{{ADDIN_" not in rendered


def test_bundled_sidebar_files_match_sha256sums() -> None:
    sidebar = desk_loader.PLUGIN_DIR / "sidebar"
    sums = sidebar / "SHA256SUMS"
    assert sums.is_file()
    entries = {}
    for line in sums.read_text(encoding="utf-8").splitlines():
        digest, name = line.split("  ", 1)
        entries[name] = digest
    assert "taskpane.html" in entries
    assert all((sidebar / name).is_file() for name in entries)
    import hashlib
    assert all(hashlib.sha256((sidebar / name).read_bytes()).hexdigest() == digest
               for name, digest in entries.items())


# ------------------------------------------------------------------------------------------------
# approval: required, deny, allow, remembered, revoke
# ------------------------------------------------------------------------------------------------

def test_session_requires_an_approved_grant(tmp_path: Path) -> None:
    with running(make_runtime(tmp_path)) as rt:
        for grant in (None, "", "not-a-grant", "x" * 300):
            status, _, body = new_session(rt, grant)
            assert status == 403 and body["approval"] == "required", body


def test_deny_path(tmp_path: Path) -> None:
    with running(make_runtime(tmp_path, choice="deny")) as rt:
        status, _, created = request(rt, "POST", f"/{APP}/connect", {"client": {"host": APP}})
        assert status == 201 and created["command"] == f"/{desk_app.COMMAND}"
        assert "Denied" in rt.slash("")  # the human picks Deny in Hermes
        status, _, polled = request(rt, "POST", f"/{APP}/connect/status",
                                    {"requestId": created["requestId"], "pollSecret": created["pollSecret"]})
        assert polled["status"] == "denied" and "grant" not in polled
        assert rt.store.status()["approved"] is False and rt.store.pending() == []
        # Asking again after a denial yields nothing either.
        status, _, again = request(rt, "POST", f"/{APP}/connect/status",
                                   {"requestId": created["requestId"], "pollSecret": created["pollSecret"]})
        assert again["status"] == "expired"


def test_no_answer_leaves_request_pending(tmp_path: Path) -> None:
    with running(make_runtime(tmp_path, choice=None)) as rt:
        out = connect(rt, decision=None)
        assert out["status"] == "pending"
        assert out["code"] in rt.slash("")  # tells the user how to answer instead


def test_poll_secret_and_origin_bind_the_request(tmp_path: Path) -> None:
    with running(make_runtime(tmp_path)) as rt:
        _, _, created = request(rt, "POST", f"/{APP}/connect", {"client": {}})
        rt.store.decide(True, created["code"])
        status, _, _ = request(rt, "POST", f"/{APP}/connect/status",
                               {"requestId": created["requestId"], "pollSecret": "wrong"})
        assert status == 401
        assert rt.store.status()["grants"] == 0  # nothing minted for a wrong secret


def test_allow_mints_a_grant_once_and_is_remembered_across_restart(tmp_path: Path) -> None:
    with running(make_runtime(tmp_path, choice="allow")) as rt:
        _, _, created = request(rt, "POST", f"/{APP}/connect", {"client": {"host": APP}})
        assert "Allowed" in rt.slash("")
        status, _, polled = request(rt, "POST", f"/{APP}/connect/status",
                                    {"requestId": created["requestId"], "pollSecret": created["pollSecret"]})
        grant = polled["grant"]
        assert polled["status"] == "approved" and len(grant) >= 40
        status, _, again = request(rt, "POST", f"/{APP}/connect/status",
                                   {"requestId": created["requestId"], "pollSecret": created["pollSecret"]})
        assert again["status"] == "expired" and "grant" not in again  # returned exactly once
        status, _, session = new_session(rt, grant)
        assert status == 201 and session["sessionId"]
    # Hermes restarts: a new runtime over the same plugin data folder; Office restarts: new pane id.
    with running(make_runtime(tmp_path)) as rt2:
        status, _, session = new_session(rt2, grant)
        assert status == 201, session
        assert rt2.store.status() == {"approved": True, "grants": 1, "pending": 0}


def test_grant_is_bound_to_the_configured_origin(tmp_path: Path) -> None:
    with running(make_runtime(tmp_path)) as rt:
        grant = connect(rt)["grant"]
    other = DeskRuntime(appcfg=_AppCfg(_free_ports(2)), data_dir=tmp_path,
                        addin_base="https://other.example.test/x/", api_server_url="http://127.0.0.1:9",
                        key_provider=KeyProvider(lambda: API_KEY), in_gateway=lambda: False)
    with running(other):
        status, _, _ = request(other, "POST", f"/{APP}/session",
                               {"protocolVersion": "1.0", "documentId": "d", "paneId": str(uuid.uuid4())},
                               headers={GRANT_HEADER: grant}, origin="https://other.example.test")
        assert status == 403


def test_revoke_closes_open_sessions_and_forgets_the_approval(tmp_path: Path) -> None:
    with running(make_runtime(tmp_path)) as rt:
        grant = connect(rt)["grant"]
        pane = str(uuid.uuid4())
        _, _, created = new_session(rt, grant, pane)
        hdrs = session_headers(created, pane)
        status, _, _ = request(rt, "POST", f"/{APP}/heartbeat", {}, headers=hdrs)
        assert status in (200, 204)
        text = rt.slash("revoke")
        assert "Revoked" in text and "1 remembered approval" in text
        status, _, _ = request(rt, "POST", f"/{APP}/heartbeat", {}, headers=hdrs)
        assert status == 401
        status, _, body = new_session(rt, grant)
        assert status == 403 and body["approval"] == "required"
        assert rt.store.status() == {"approved": False, "grants": 0, "pending": 0}


def test_revoke_shows_removal_steps_for_this_install_ca_and_changes_no_trust(tmp_path: Path, capsys) -> None:
    from types import SimpleNamespace
    rt, other = make_runtime(tmp_path / "a"), make_runtime(tmp_path / "b")
    assert "Stop the bridge" not in rt.slash("revoke")  # no local CA yet: nothing to remove
    rt.bridge.tls.ensure()
    other.bridge.tls.ensure()
    steps = rt.bridge.tls.removal_instructions()
    assert rt.bridge.tls.ca_thumbprint() in steps and other.bridge.tls.ca_thumbprint() not in steps
    assert rt.slash("revoke").endswith("\n" + steps)
    assert rt.cli_run(SimpleNamespace(action="revoke", code=None)) == 0
    assert capsys.readouterr().out.endswith("\n" + steps + "\n")
    assert rt.bridge.tls.ca_path.is_file() and rt.bridge.port is None


def test_revoke_from_another_process_reaches_the_running_bridge(tmp_path: Path) -> None:
    with running(make_runtime(tmp_path)) as rt:
        grant = connect(rt)["grant"]
        pane = str(uuid.uuid4())
        _, _, created = new_session(rt, grant, pane)
        time.sleep(0.01)
        ApprovalStore(tmp_path, APP).revoke()  # e.g. `hermes sidecar-<app> revoke` in a terminal
        status, _, _ = request(rt, "POST", f"/{APP}/heartbeat", {}, headers=session_headers(created, pane))
        assert status == 401


def test_cli_allow_and_status(tmp_path: Path, capsys) -> None:
    with running(make_runtime(tmp_path)) as rt:
        _, _, created = request(rt, "POST", f"/{APP}/connect", {"client": {}})
        cli = make_runtime(tmp_path)  # a separate process's view (no bridge running there)

        class Args:
            action = "allow"
            code = created["code"]

        assert cli.cli_run(Args) == 0
        _, _, polled = request(rt, "POST", f"/{APP}/connect/status",
                               {"requestId": created["requestId"], "pollSecret": created["pollSecret"]})
        assert polled["status"] == "approved"
        Args.action = "status"
        assert cli.cli_run(Args) == 0
        assert "allowed (1 sidebar grant" in capsys.readouterr().out


def test_no_tool_can_approve(tmp_path: Path) -> None:
    schemas = json.loads((desk_loader.PLUGIN_DIR / "tool_schemas.json").read_text(encoding="utf-8"))
    for s in schemas:
        assert not any(w in s["name"] for w in ("approve", "allow", "grant", "connect", "revoke"))


# ------------------------------------------------------------------------------------------------
# chat proxy, key handling, no key file, no key in logs
# ------------------------------------------------------------------------------------------------

def test_chat_proxy_forwards_with_key_and_forced_session(tmp_path: Path) -> None:
    upstream = Upstream()
    try:
        with running(make_runtime(tmp_path, api_url=upstream.url)) as rt:
            grant = connect(rt)["grant"]
            pane = str(uuid.uuid4())
            _, _, created = new_session(rt, grant, pane)
            hdrs = session_headers(created, pane)
            status, headers, body = request(rt, "POST", "/v1/chat/completions",
                                            {"model": "hermes-agent", "messages": [], "stream": True},
                                            headers={**hdrs, "Authorization": "Bearer pane-supplied"})
            assert status == 200 and headers["Access-Control-Allow-Origin"] == ORIGIN
            assert b"[DONE]" in body
            seen = upstream.requests[-1]
            assert seen["path"] == "/v1/chat/completions"
            assert seen["headers"]["Authorization"] == f"Bearer {API_KEY}"
            assert seen["headers"]["X-Hermes-Session-Id"] == created["sessionId"]
            assert "pane-supplied" not in json.dumps(seen["headers"])
            assert API_KEY not in json.dumps(headers) and API_KEY.encode() not in body
            # Without a session the proxy refuses and forwards nothing.
            before = len(upstream.requests)
            status, _, _ = request(rt, "POST", "/v1/chat/completions", {"messages": []})
            assert status == 401 and len(upstream.requests) == before
    finally:
        upstream.close()


def test_chat_proxy_without_key_or_upstream_fails_cleanly(tmp_path: Path) -> None:
    with running(make_runtime(tmp_path, key=None)) as rt:
        grant = connect(rt)["grant"]
        pane = str(uuid.uuid4())
        _, _, created = new_session(rt, grant, pane)
        status, _, body = request(rt, "POST", "/v1/chat/completions", {"messages": []},
                                  headers=session_headers(created, pane))
        assert status == 503 and "API_SERVER_KEY" in body["error"]
    closed = _free_ports(1)[0]
    with running(make_runtime(tmp_path, api_url=f"http://127.0.0.1:{closed}")) as rt:
        pane = str(uuid.uuid4())
        _, _, created = new_session(rt, grant, pane)
        status, _, body = request(rt, "POST", "/v1/chat/completions", {"messages": []},
                                  headers=session_headers(created, pane))
        assert status == 502 and API_KEY not in json.dumps(body)


def test_api_key_is_never_logged_or_written(tmp_path: Path, caplog) -> None:
    home = tmp_path / "home"
    data = home / "plugin-data" / desk_app.PLUGIN_ID
    caplog.set_level(logging.DEBUG)
    err = io.StringIO()
    upstream = Upstream(status=401)
    try:
        with redirect_stderr(err), running(make_runtime(data, api_url=upstream.url)) as rt:
            grant = connect(rt)["grant"]
            pane = str(uuid.uuid4())
            _, _, created = new_session(rt, grant, pane)
            request(rt, "POST", "/v1/chat/completions", {"messages": []}, headers=session_headers(created, pane))
            rt.slash("status")
            rt.slash("revoke")
            text = repr(rt) + repr(rt.key_provider) + repr(rt.bridge) + rt.status_text()
    finally:
        upstream.close()
    logged = "\n".join(r.getMessage() for r in caplog.records) + err.getvalue()
    assert API_KEY not in logged and API_KEY not in text
    assert grant not in logged
    files = [p for p in home.rglob("*") if p.is_file()]
    assert files, "the approval store should have written its state files"
    for p in files:
        raw = p.read_bytes()
        assert API_KEY.encode() not in raw, p
        assert grant.encode() not in raw, p  # only sha256(grant) is stored
        assert not p.name.endswith(".key")
        if os.name != "nt":
            assert (p.stat().st_mode & 0o077) == 0, p
    assert {p.name for p in files} <= {"approvals.json", "pending.json", "ca.pem", "ca-key.pem", "localhost.pem", "localhost-key.pem"}


def test_plugin_never_reads_or_writes_the_old_key_files() -> None:
    sources = "\n".join(p.read_text(encoding="utf-8") for p in desk_loader.PLUGIN_DIR.rglob("*.py") if p.name != "desk_tls.py")
    for needle in ("office-bridge", ".key\"", "keyFile", "read_bytes(", ".env\"", "auth.json", "os.environ.get(\"API"):
        assert needle not in sources, needle
    assert "get_secret(\"API_SERVER_KEY\")" in sources


# ------------------------------------------------------------------------------------------------
# plugin wiring: autostart only in the gateway, tool calls reach the paired pane
# ------------------------------------------------------------------------------------------------

class FakeCtx:
    def __init__(self, settings: dict) -> None:
        self.settings = settings
        self.tools: dict = {}
        self.commands: dict = {}
        self.cli: dict = {}
        self.unload: list = []

    def get_config(self, key, default=None):
        return self.settings.get(key, default)

    def register_tool(self, **kw):
        self.tools[kw["name"]] = kw

    def register_command(self, name, handler, **kw):
        self.commands[name] = handler

    def register_cli_command(self, **kw):
        self.cli[kw["name"]] = kw

    def on_unload(self, cb):
        self.unload.append(cb)

    def ask_choice(self, question, choices):
        return "allow"


@pytest.fixture
def plugin_env(tmp_path: Path, monkeypatch):
    data = tmp_path / "data"
    module = desk_loader.fresh("desk_plugin_wiring")
    monkeypatch.setattr(module.desk_bridge, "_data_dir_for", lambda plugin_id: data)
    monkeypatch.setattr(module.desk_bridge, "_read_api_server_key", lambda: API_KEY)
    yield module
    rt = getattr(module, "_RUNTIME", None)
    if rt is not None:
        rt.stop()


def test_register_does_not_listen_outside_the_gateway(plugin_env, monkeypatch) -> None:
    monkeypatch.delenv("_HERMES_GATEWAY", raising=False)
    ctx = FakeCtx({"addin_base": ADDIN_BASE})
    plugin_env.register(ctx)
    assert plugin_env._RUNTIME.bridge.port is None
    assert desk_app.COMMAND in ctx.commands and desk_app.COMMAND in ctx.cli and ctx.unload


def test_register_starts_bridge_in_gateway_and_tool_reaches_pane(plugin_env, monkeypatch) -> None:
    monkeypatch.setenv("_HERMES_GATEWAY", "1")
    monkeypatch.setattr(plugin_env.desk_app, "PORTS", _free_ports(2))
    ctx = FakeCtx({"addin_base": ADDIN_BASE})
    plugin_env.register(ctx)
    rt = plugin_env._RUNTIME
    assert rt.bridge.port in plugin_env.desk_app.PORTS
    grant = connect(rt)["grant"]
    pane = str(uuid.uuid4())
    _, _, created = new_session(rt, grant, pane)
    hdrs = session_headers(created, pane)
    schemas = json.loads((desk_loader.PLUGIN_DIR / "tool_schemas.json").read_text(encoding="utf-8"))
    read_tool = next(s["name"] for s in schemas if "_get_" in s["name"] or "_read_" in s["name"])
    canonical = read_tool[len(desk_app.TOOL_PREFIX):]
    seen: dict = {}

    def pane_loop():
        deadline = time.time() + 10
        while time.time() < deadline:
            status, _, ticket = request(rt, "GET", f"/{APP}/pending", headers=hdrs)
            if status == 200:
                seen.update(ticket)
                request(rt, "POST", f"/{APP}/result", {"ticketId": ticket["ticketId"], "result": {"ok": True, "v": 1}},
                        headers=hdrs)
                return
            time.sleep(0.05)

    t = threading.Thread(target=pane_loop, daemon=True)
    t.start()
    out = ctx.tools[read_tool]["handler"]({}, session_id=created["sessionId"], user_task="read it", turn_id="t1")
    t.join(5)
    assert seen.get("tool") == canonical
    assert json.loads(out) == {"ok": True, "v": 1}
    # Another chat (not the sidebar's session) cannot drive the pane.
    foreign = json.loads(ctx.tools[read_tool]["handler"]({}, session_id=str(uuid.uuid4()), user_task="x", turn_id="t"))
    assert foreign["ok"] is False and foreign["error"]["category"] == "ForeignSession"


def test_revoked_session_blocks_tool_ticket_immediately(plugin_env, monkeypatch):
    monkeypatch.delenv("_HERMES_GATEWAY", raising=False)
    monkeypatch.setattr(plugin_env.desk_app, "PORTS", _free_ports(2))
    ctx = FakeCtx({"addin_base": ADDIN_BASE})
    plugin_env.register(ctx)
    rt = plugin_env._RUNTIME
    with running(rt):
        grant = connect(rt)["grant"]
        status, _, created = new_session(rt, grant, str(uuid.uuid4()))
        assert status == 201 and isinstance(created, dict)
        session = created["sessionId"]
        assert rt.bridge.has_session(session) and rt.bridge.session_allowed(session)
        # A separate CLI process removes persisted grants, not this bridge's sessions.
        other_process_store = ApprovalStore(rt.store.data_dir, APP)
        assert other_process_store.revoke() == 1
        assert rt.bridge.has_session(session) and not rt.bridge.session_allowed(session)
        sent = []
        monkeypatch.setattr(plugin_env, "_invoke_bridge_http", lambda *a, **kw: sent.append(a) or {"ok": True})
        name = desk_app.TOOL_PREFIX + "read_range"
        out = json.loads(ctx.tools[name]["handler"]({}, session_id=session))
        assert out.get("ok") is False, "revoked session must not send a tool ticket"
        assert out["error"]["category"] == "ForeignSession"
        assert sent == []


def test_allow_question_discloses_full_agent_authority(tmp_path):
    rt = make_runtime(tmp_path)
    questions = []
    rt._ask_choice = lambda q, c: questions.append(q) or "deny"
    with running(rt):
        request(rt, "POST", f"/{APP}/connect", {})
        rt.slash("")
    assert len(questions) == 1
    assert f"site at {ORIGIN}" in questions[0]
    assert "your Hermes agent and all its tools until you revoke" in questions[0]
    assert "Use a site you control or trust; open only documents you trust" in questions[0]


def test_installed_license_and_notice_match_root():
    for name in ("LICENSE", "NOTICE"):
        assert (desk_loader.PLUGIN_DIR / name).read_bytes() == (desk_loader.PLUGIN_DIR.parents[1] / name).read_bytes()


def test_invalid_addin_base_disables_runtime_but_still_registers_tools(plugin_env, monkeypatch) -> None:
    monkeypatch.delenv("_HERMES_GATEWAY", raising=False)
    ctx = FakeCtx({"addin_base": "http://insecure.example/"})
    plugin_env.register(ctx)
    assert plugin_env._RUNTIME is None and ctx.tools
    name = next(iter(ctx.tools))
    out = json.loads(ctx.tools[name]["handler"]({}, session_id=str(uuid.uuid4()), user_task="read", turn_id="t"))
    assert out["ok"] is False and out["error"]["category"] in ("Configuration", "Policy")


def test_cli_manifest_uses_running_gateway_bridge_without_starting_one(tmp_path: Path, monkeypatch, capsys) -> None:
    """A separate `hermes <command> manifest` process must point at the gateway's bridge, not bind its own."""
    rt = DeskRuntime(appcfg=_AppCfg(_free_ports(2)), data_dir=tmp_path,
                     addin_base=desk_app.DEFAULT_ADDIN_BASE, api_server_url="http://127.0.0.1:9",
                     key_provider=KeyProvider(lambda: API_KEY), in_gateway=lambda: False)
    gateway_port = rt.cfg.PORTS[1]
    monkeypatch.setattr(rt, "_find_running_default_port", lambda: gateway_port)

    class Args:
        action = "manifest"
        code = None

    try:
        assert rt.cli_run(Args) == 0
        assert rt.bridge.port is None  # no throwaway listener on the next free port
        manifest = (tmp_path / "manifest.xml").read_text(encoding="utf-8")
        assert f"https://localhost:{gateway_port}/{desk_app.COMMAND}/taskpane.html" in manifest
        assert f"<AppDomain>https://localhost:{gateway_port}</AppDomain>" in manifest

        monkeypatch.setattr(rt, "_find_running_default_port", lambda: None)
        assert "Start the Hermes gateway" in rt.slash("manifest")
        assert rt.bridge.port is None
    finally:
        rt.stop()


def test_same_origin_get_without_origin_header_is_accepted_only_for_the_plugin_served_page(tmp_path: Path) -> None:
    """Same-origin GETs carry no Origin header; the plugin-served page must still discover its bridge."""
    rt = DeskRuntime(appcfg=_AppCfg(_free_ports(2)), data_dir=tmp_path / "default",
                     addin_base=desk_app.DEFAULT_ADDIN_BASE, api_server_url="http://127.0.0.1:9",
                     key_provider=KeyProvider(lambda: API_KEY), in_gateway=lambda: False)
    with running(rt):
        status, _, body = request(rt, "GET", f"/{APP}/hello", origin=None)
        assert status == 200 and body["app"] == APP
        status, _, _ = request(rt, "GET", f"/{APP}/hello", origin=None,
                               headers={"Sec-Fetch-Site": "same-origin"})
        assert status == 200
        for site in ("cross-site", "same-site", "none"):
            status, _, _ = request(rt, "GET", f"/{APP}/hello", origin=None, headers={"Sec-Fetch-Site": site})
            assert status == 403, site
        status, _, _ = request(rt, "GET", f"/{APP}/hello", origin="https://evil.example")
        assert status == 403
        # Writes still need an explicit, matching Origin header.
        status, _, _ = request(rt, "POST", f"/{APP}/connect", body={}, origin=None)
        assert status == 403
    # A self-hosted override keeps the strict exact-origin rule: no Origin, no access.
    with running(make_runtime(tmp_path / "override")) as rt:
        status, _, _ = request(rt, "GET", f"/{APP}/hello", origin=None)
        assert status == 403
