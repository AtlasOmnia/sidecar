"""Exercise the real TLS Desk/companion composition, not a stand-alone proxy."""
import concurrent.futures
import threading
import time
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from test_desk_bridge_security import (
    APP, connect, make_runtime, new_session, request, running, session_headers,
)


def test_chat_survives_delayed_sse_and_concurrent_tool_roundtrip(tmp_path):
    seen = []
    started = threading.Event()
    chunks = [b'data: {"choices":[{"delta":{"content":"Edited"}}]}\n\n',
              b'data: {"choices":[{"delta":{"content":" successfully"}}]}\n\ndata: [DONE]\n\n']

    class Upstream(BaseHTTPRequestHandler):
        def log_message(self, format, *args):
            pass

        def do_POST(self):
            seen.append((self.path, self.rfile.read(int(self.headers["Content-Length"]))))
            self.send_response(200)
            self.send_header("Content-Type", "text/event-stream")
            self.send_header("Content-Length", str(sum(map(len, chunks))))
            self.end_headers()
            self.wfile.flush()
            started.set()
            # Both time-to-first-event and inter-event silence exceed the
            # companion's one-second downstream *write* budget.
            for chunk in chunks:
                time.sleep(3)
                try:
                    self.wfile.write(chunk)
                    self.wfile.flush()
                except OSError:
                    return

    upstream = ThreadingHTTPServer(("127.0.0.1", 0), Upstream)
    thread = threading.Thread(target=upstream.serve_forever, daemon=True)
    thread.start()
    try:
        with running(make_runtime(tmp_path, api_url=f"http://127.0.0.1:{upstream.server_port}")) as rt:
            grant = connect(rt)["grant"]
            pane = str(uuid.uuid4())
            status, _, session = new_session(rt, grant, pane)
            assert status == 201
            assert isinstance(session, dict)
            headers = session_headers(session, pane)
            with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
                chat = pool.submit(request, rt, "POST", "/v1/chat/completions",
                                   {"model": "hermes-agent", "messages": [], "stream": True}, headers)
                assert started.wait(5)
                invoke = pool.submit(request, rt, "POST", f"/{APP}/invoke", {
                    "protocolVersion": "1.0", "sessionId": session["sessionId"],
                    "requestId": str(uuid.uuid4()), "tool": "get_presentation_overview", "arguments": {},
                }, {"X-Hermes-Plugin-Key": rt.bridge.plugin_key})
                deadline = time.monotonic() + 2
                while True:
                    status, _, ticket = request(rt, "GET", f"/{APP}/pending", headers=headers)
                    if status == 200:
                        break
                    assert time.monotonic() < deadline
                    time.sleep(0.02)
                assert isinstance(ticket, dict)
                assert ticket["tool"] == "get_presentation_overview"
                status, _, _ = request(rt, "POST", f"/{APP}/result", {
                    "ticketId": ticket["ticketId"], "result": {"ok": True, "slideCount": 5},
                }, headers=headers)
                assert status == 200
                assert invoke.result(timeout=5)[0] == 200
                status, _, body = chat.result(timeout=10)
                assert status == 200
                assert body == b"".join(chunks)
                assert len(seen) == 1
                assert seen[0][0] == "/v1/chat/completions"
    finally:
        upstream.shutdown()
        upstream.server_close()
        thread.join(5)
