"""Runtime registration and dispatch tests for the Word Hermes plugin."""
from __future__ import annotations

import importlib.util
import json
import os
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO))
PLUGIN_DIR = REPO / "plugins" / "atlas-sidecar-word"


class FakeContext:
    def __init__(self) -> None:
        self.tools: list[dict] = []

    def register_tool(self, **kwargs) -> None:
        self.tools.append(kwargs)


def load_plugin():
    import desk_loader
    return desk_loader.fresh("desk_plugin_runtime_test")


def test_registers_every_generated_schema_without_core_override() -> None:
    module = load_plugin()
    context = FakeContext()
    module.register(context)
    schemas = json.loads((PLUGIN_DIR / "tool_schemas.json").read_text(encoding="utf-8"))
    assert len(context.tools) == len(schemas) == 12
    assert {tool["name"] for tool in context.tools} == {schema["name"] for schema in schemas}
    assert all(tool["toolset"] == "sidecar-word" for tool in context.tools)
    assert all(tool["name"].startswith("sidecar_word_") for tool in context.tools)
    assert all(tool.get("override", False) is False for tool in context.tools)
    assert all(tool.get("check_fn") is None for tool in context.tools)


def test_handler_rejects_model_supplied_bridge_session() -> None:
    module = load_plugin()
    calls: list[dict] = []

    def fake_invoke(session_id: str, tool: str, arguments: dict) -> dict:
        calls.append({"session_id": session_id, "tool": tool, "arguments": arguments})
        return {"ok": True, "data": {"cells": [[1]]}}

    setattr(module, "invoke_bridge", fake_invoke)
    context = FakeContext()
    module.register(context)
    read_tool = next(tool for tool in context.tools if tool["name"] == "sidecar_word_read_document_text")
    result = json.loads(read_tool["handler"](
        {
            "bridgeSessionId": "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
        },
        session_id="bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb",
    ))
    assert result == {
        "ok": False,
        "error": {
            "category": "Validation",
            "message": "bridgeSessionId is trusted request context and cannot be supplied by the model",
        },
    }
    assert calls == []


def test_handler_uses_trusted_session_id_kwarg() -> None:
    module = load_plugin()
    calls: list[dict] = []

    def fake_invoke(session_id: str, tool: str, arguments: dict) -> dict:
        calls.append({"session_id": session_id, "tool": tool, "arguments": arguments})
        return {"ok": True, "data": {"cells": [[1]]}}

    setattr(module, "invoke_bridge", fake_invoke)
    context = FakeContext()
    module.register(context)
    read_tool = next(tool for tool in context.tools if tool["name"] == "sidecar_word_read_document_text")
    result = json.loads(read_tool["handler"](
        {},
        session_id="aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
    ))
    assert result == {"cells": [[1]]}
    assert calls == [{
        "session_id": "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
        "tool": "read_document_text",
        "arguments": {},
    }]


def test_client_rejects_extra_or_mismatched_success_envelopes() -> None:
    import http.server
    import tempfile
    import threading

    module = load_plugin()
    bridge_client = importlib.import_module(module.bridge_client.__name__)
    ticket_id = "550e8400-e29b-41d4-a716-446655440000"
    key = "envelope-test-key-" + ("e" * 32)

    class Handler(http.server.BaseHTTPRequestHandler):
        mode = "extra"
        hits = 0

        def log_message(self, format: str, *_args: object) -> None:
            return

        def do_POST(self) -> None:  # noqa: N802
            length = int(self.headers["Content-Length"])
            request = json.loads(self.rfile.read(length))
            type(self).hits += 1
            response = {
                "ok": True,
                "ticketId": ticket_id,
                "requestId": request["requestId"],
                "data": {"cells": [[1]]},
            }
            if type(self).mode == "extra":
                response["unexpected"] = True
            else:
                response["requestId"] = "660e8400-e29b-41d4-a716-446655440001"
            body = json.dumps(response).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

    for mode in ("extra", "mismatch"):
        server = http.server.ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        Handler.mode = mode
        Handler.hits = 0
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        try:
            try:
                bridge_client.invoke_bridge(
                    "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa",
                    "read_document_text",
                    {},
                    base_url=f"http://127.0.0.1:{server.server_address[1]}",
                    key=key,
                )
            except bridge_client.BridgeError as exc:
                assert "response is invalid" in str(exc)
            else:
                raise AssertionError("malformed companion response was accepted")
            assert Handler.hits == 1
        finally:
            server.shutdown()
            server.server_close()
            thread.join(2)


def test_client_requires_running_bridge_and_in_memory_key() -> None:
    module = load_plugin()
    bridge_client = importlib.import_module(module.bridge_client.__name__)
    sid = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"
    for base_url, key in ((None, "k" * 64), ("http://127.0.0.1:47630", ""), ("http://10.0.0.2:47630", "k" * 64)):
        try:
            bridge_client.invoke_bridge(sid, "read_document_text", {}, base_url=base_url, key=key)
        except bridge_client.BridgeError as exc:
            assert exc.category in ("Unavailable", "Configuration")
        else:
            raise AssertionError("invoke without a running local bridge was accepted")


if __name__ == "__main__":
    tests = [(name, value) for name, value in list(globals().items())
             if name.startswith("test_") and callable(value)]
    passed = 0
    for name, test in tests:
        test()
        passed += 1
        print(f"  PASS {name}")
    print(f"\n{passed} passed, 0 failed, {len(tests)} total")
