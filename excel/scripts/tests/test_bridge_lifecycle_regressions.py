"""Live sockets and real plugin registration; no CLI invocation or OS trust writes."""
import http.client
import sys
import ssl
from types import SimpleNamespace

import pytest

from test_desk_bridge_security import (
    APP, ORIGIN, ADDIN_BASE, FakeCtx, _free_ports, make_runtime, plugin_env,
)


def test_stop_closes_existing_tls_keepalive(tmp_path):
    rt = make_runtime(tmp_path)
    rt.ensure_started()
    conn = http.client.HTTPSConnection("localhost", rt.bridge.port, timeout=2,
                                       context=rt.bridge.tls.client_context())
    try:
        headers = {"Host": f"localhost:{rt.bridge.port}", "Origin": ORIGIN}
        conn.request("GET", f"/{APP}/hello", headers=headers)
        response = conn.getresponse()
        response.read()
        assert response.status == 200
        assert not response.will_close
        original_socket = conn.sock
        assert original_socket is not None
        rt.stop()
        # Reuse the SAME socket, not HTTPSConnection's automatic reconnect path.
        try:
            original_socket.sendall((f"GET /{APP}/hello HTTP/1.1\r\n"
                                     f"Host: {headers['Host']}\r\nOrigin: {ORIGIN}\r\n\r\n").encode())
            received = original_socket.recv(4096)
        except (ConnectionError, ssl.SSLEOFError):
            received = b""  # Only EOF/reset counts; a read timeout must fail.
        assert received == b"", "stopped bridge returned an HTTP response on its old TLS connection"
    finally:
        conn.close()
        rt.stop()


@pytest.mark.parametrize("action", ["status", "allow", "revoke"])
@pytest.mark.parametrize("gateway_marker", ["0", "1"])
def test_cli_registration_ignores_inherited_gateway_marker(plugin_env, monkeypatch, capsys, action, gateway_marker):
    monkeypatch.setenv("_HERMES_GATEWAY", gateway_marker)
    monkeypatch.setattr(sys, "argv", ["hermes", "-p", "test", f"sidecar-{APP}", action])
    monkeypatch.setattr(plugin_env.desk_app, "PORTS", _free_ports(2))
    ctx = FakeCtx({"addin_base": ADDIN_BASE})
    plugin_env.register(ctx)
    rt = plugin_env._RUNTIME
    assert rt.bridge.port is None, "short-lived CLI registration started a bridge listener"
    rt.bridge.tls.ensure()  # Existing identity must not cause status to print trust instructions.
    for _ in range(2):
        rt.cli_run(SimpleNamespace(action=action, code=None))
    output = capsys.readouterr().out
    assert "add-trusted-cert" not in output
    assert "Manual TLS trust" not in output
    assert rt.cli_run(SimpleNamespace(action="trust", code=None)) == 0
    assert "Manual TLS trust" in capsys.readouterr().out
    assert rt.bridge.port is None
