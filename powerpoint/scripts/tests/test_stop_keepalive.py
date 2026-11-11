"""Stopping the bridge must retire handlers as well as the listening socket."""
import http.client

from test_desk_bridge_security import APP, ORIGIN, make_runtime, running


def test_stop_closes_existing_tls_keepalive_connection(tmp_path):
    with running(make_runtime(tmp_path)) as rt:
        conn = http.client.HTTPSConnection("localhost", rt.bridge.port,
                                          context=rt.bridge.tls.client_context(), timeout=3)
        try:
            headers = {"Origin": ORIGIN, "Connection": "keep-alive"}
            conn.request("GET", f"/{APP}/hello", headers=headers)
            response = conn.getresponse()
            assert response.status == 200
            response.read()
            original_socket = conn.sock
            assert original_socket is not None
            rt.bridge.stop()
            # Raw reuse prevents HTTPSConnection from transparently reconnecting.
            try:
                original_socket.sendall((f"GET /{APP}/hello HTTP/1.1\r\n"
                                         f"Host: localhost:{original_socket.getpeername()[1]}\r\n"
                                         f"Origin: {ORIGIN}\r\n\r\n").encode())
                received = original_socket.recv(4096)
            except TimeoutError:
                raise AssertionError("stopped bridge left the TLS connection open")
            except OSError:
                received = b""
            assert received == b"", "stopped bridge served an HTTP response on an old TLS connection"
        finally:
            conn.close()
