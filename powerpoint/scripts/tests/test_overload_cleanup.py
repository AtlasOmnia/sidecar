"""Offline overload/response regressions with bounded fake and loopback sockets."""
from __future__ import annotations
import http.client
import importlib.util
import io
from pathlib import Path
import socket
import sys
import threading
import time
from unittest.mock import Mock
import pytest

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from desk_plugin.companion.handler import CompanionHandler
from desk_plugin.companion.server import PowerPointCompanionServer, _BoundedThreadingHTTPServer


class SlowSocket:
    def __init__(self):
        self.sent = []
        self.threads = []
        self.closed = False
        self.release = threading.Event()
        self.reads = 0
    def settimeout(self, timeout):
        self.timeout = timeout
    def recv(self, size):
        self.threads.append(threading.get_ident())
        self.release.wait(0.15)
        self.reads += size
        return b""
    def sendall(self, data):
        self.threads.append(threading.get_ident())
        self.sent.append(data)
    def shutdown(self, how):
        if how == socket.SHUT_RDWR:
            self.release.set()
    def close(self):
        self.closed = True
        self.release.set()


@pytest.mark.parametrize("kind", ["companion"])
def test_overload_never_writes_or_reads_on_accept_thread(kind):
    server = (proxy.create_server('127.0.0.1', 0, max_workers=1) if kind=='proxy'
              else _BoundedThreadingHTTPServer(('127.0.0.1',0), CompanionHandler, max_handler_threads=1))
    sem = server._worker_slots if kind=='proxy' else server._handler_semaphore
    assert sem.acquire(False)
    request = SlowSocket()
    caller = threading.get_ident()
    try:
        server.process_request(request, ('127.0.0.1', 1))
        time.sleep(0.02)
        assert caller not in request.threads
        assert len(server._overload_cleanup_threads) == 2
    finally:
        sem.release()
        server.server_close()
    assert request.closed
    assert all(not thread.is_alive() for thread in server._overload_cleanup_threads)


@pytest.mark.parametrize("kind", ["companion"])
def test_overload_fixed_queue_refuses_before_http_when_full(kind):
    server = (proxy.create_server('127.0.0.1',0,max_workers=1) if kind=='proxy'
              else _BoundedThreadingHTTPServer(('127.0.0.1',0),CompanionHandler,max_handler_threads=1))
    sem = server._worker_slots if kind=='proxy' else server._handler_semaphore
    sem.acquire(False)
    requests = [SlowSocket() for _ in range(20)]
    try:
        for request in requests:
            server.process_request(request, ('127.0.0.1',1))
        assert len(server._overload_cleanup_sockets) <= 10
        assert server._overload_cleanup_queue.qsize() <= 8
        refused = [r for r in requests if r.closed]
        assert refused and all(not r.sent for r in refused)
    finally:
        sem.release()
        server.server_close()
    assert all(r.closed for r in requests)




@pytest.mark.parametrize("kind", ["companion"])
def test_cleanup_constructor_failure_releases_listener(kind, monkeypatch):
    real_start = threading.Thread.start
    started = []
    def fail_second(self):
        if started:
            raise RuntimeError('synthetic cleanup startup failure')
        started.append(self)
        return real_start(self)
    with socket.socket() as probe:
        probe.bind(('127.0.0.1',0))
        address = probe.getsockname()
    monkeypatch.setattr(threading.Thread, 'start', fail_second)
    with pytest.raises(RuntimeError, match='cleanup'):
        if kind=='proxy':
            proxy.create_server(*address)
        else:
            _BoundedThreadingHTTPServer(address, CompanionHandler)
    assert started and all(not thread.is_alive() for thread in started)
    with socket.socket() as probe:
        probe.bind(address)


def test_pending_poll_always_closes_framed_reply():
    server = PowerPointCompanionServer(port=0,plugin_key='synthetic')
    session = server.sessions.create('deck','pane','key')
    server.start()
    connection = http.client.HTTPConnection(*server.server_address(), timeout=2)
    try:
        connection.request('GET','/powerpoint/pending',headers={'Origin':'https://localhost:3002','X-Hermes-Session-Id':session.session_id,'X-Hermes-Session-Key':'key','X-Hermes-Pane-Id':'pane'})
        response = connection.getresponse()
        assert response.status==204
        assert response.getheader('Connection')=='close'
        assert response.read()==b''
    finally:
        connection.close()
        server.stop()


def test_response_one_absolute_budget_covers_headers_body_and_flush(monkeypatch):
    from desk_plugin.companion import handler as module
    clock = [100.0]
    monkeypatch.setattr(module.time, 'monotonic', lambda: clock[0])
    handler = object.__new__(CompanionHandler)
    handler.command='GET'
    handler.path='/powerpoint/health'
    handler.request_version='HTTP/1.1'
    handler.requestline='GET /powerpoint/health HTTP/1.1'
    handler.connection=Mock()
    writes=[]
    class SlowWriter(io.BytesIO):
        def write(self,data):
            writes.append(data)
            clock[0] += 0.6
            return super().write(data)
        def flush(self):
            clock[0] += 0.6
    handler.wfile=SlowWriter()
    handler.close_connection=False
    handler._send_json(200,{'ok':True},start=100.0)
    assert handler.close_connection
    handler.connection.close.assert_called_once()
    timeouts=[call.args[0] for call in handler.connection.settimeout.call_args_list]
    assert timeouts and max(timeouts)<=1.0
    assert sum(b'HTTP/1.1' in data for data in writes)==1


def test_cleanup_drain_uses_byte_cap(monkeypatch):
    from desk_plugin.companion import overload_cleanup as module
    monkeypatch.setattr(module, 'OVERLOAD_DRAIN_BYTES', 32)
    server = _BoundedThreadingHTTPServer(('127.0.0.1',0),CompanionHandler,max_handler_threads=1)
    server._handler_semaphore.acquire(False)
    request = SlowSocket()
    head = list(b'GET /overload HTTP/1.1\r\nHost: localhost\r\n\r\n')
    def recv(size):
        if head:
            return bytes([head.pop(0)])
        request.reads += size
        return b'x'*size
    request.recv = recv
    try:
        server.process_request(request,('127.0.0.1',1))
        assert request.release.wait(1)
        assert request.reads == 32
        assert request.sent == [server._overload_response]
    finally:
        server._handler_semaphore.release()
        server.server_close()


def test_connection_close_token_is_canonical_and_applies_before_write():
    handler = object.__new__(CompanionHandler)
    handler.command='GET'
    handler.path='/powerpoint/health'
    handler.request_version='HTTP/1.1'
    handler.requestline='GET /powerpoint/health HTTP/1.1'
    handler.connection=Mock()
    handler.wfile=io.BytesIO()
    handler.close_connection=False
    handler._send_error(400,'fixed',start=time.monotonic(),extra_headers={'Connection':'keep-alive','connection':'CLOSE'})
    raw = handler.wfile.getvalue()
    assert raw.count(b'Connection:') == 1
    assert b'Connection: close\r\n' in raw
    assert handler.close_connection


def test_cleanup_absolute_deadline_stops_before_byte_cap(monkeypatch):
    from desk_plugin.companion import overload_cleanup as module
    import queue
    clock=[100.0]
    monkeypatch.setattr(module.time,'monotonic',lambda:clock[0])
    owner=object.__new__(module.OverloadCleanupMixin)
    owner._overload_cleanup_queue=queue.Queue()
    owner._overload_cleanup_lock=threading.Lock()
    owner._overload_cleanup_stop=threading.Event()
    # Isolate the post-response drain budget; pre-drain has its own clock test.
    owner._drain_overload_request=lambda *args:None
    request=SlowSocket()
    owner._overload_cleanup_sockets={request}
    def sendall(data):
        clock[0] += 0.06
    def recv(size):
        clock[0] += 0.09
        request.reads += size
        return b'x'*size
    request.sendall=sendall
    request.recv=recv
    def close(request):
        request.close()
        owner._overload_cleanup_stop.set()
    owner._close_overload_request=close
    owner._overload_cleanup_queue.put(request)
    owner._overload_cleanup_loop()
    assert request.reads == 8192
    assert clock[0] > 100.2
    assert not owner._overload_cleanup_sockets
    assert owner._overload_cleanup_queue.unfinished_tasks == 0


@pytest.mark.parametrize('failure',[OSError,RuntimeError])
@pytest.mark.parametrize('expect',[False,True])
def test_route_failures_emit_final_500_after_optional_continue(failure,expect,monkeypatch):
    from desk_plugin.companion import protocol
    def fail_key():
        raise failure('synthetic key failure')
    monkeypatch.setattr(protocol,'generate_session_key',fail_key)
    server=PowerPointCompanionServer(port=0,plugin_key='synthetic')
    server.start()
    body=b'{"protocolVersion":"1.0","documentId":"deck","paneId":"pane"}'
    try:
        with socket.create_connection(server.server_address(),timeout=2) as sock:
            sock.settimeout(2)
            extra='Expect: 100-continue\r\n' if expect else ''
            sock.sendall((f'POST /powerpoint/session HTTP/1.1\r\nHost: localhost\r\nOrigin: https://localhost:3002\r\nContent-Type: application/json\r\nContent-Length: {len(body)}\r\nConnection: close\r\n'+extra+'\r\n').encode()+body)
            sock.shutdown(socket.SHUT_WR)
            reply=b''
            while True:
                chunk=sock.recv(4096)
                if not chunk:
                    break
                reply+=chunk
            assert reply.count(b'HTTP/1.1 500 ')==1
            assert reply.count(b'HTTP/1.1 100 Continue')==int(expect)
            assert b'synthetic key failure' not in reply
            assert server.sessions.active_count==0
    finally:
        server.stop()


@pytest.mark.parametrize('stale',[False,True])
@pytest.mark.parametrize('header_block',[b'X-Large: '+b'x'*65536+b'\r\n\r\n', b'X-A: x\r\n'*101+b'\r\n'],ids=['overlong','too-many'])
def test_parser_errors_do_not_require_or_reuse_cors_state(stale,header_block):
    from email.message import Message
    handler=object.__new__(CompanionHandler)
    handler.raw_requestline=b'POST /powerpoint/session HTTP/1.1\r\n'
    handler.rfile=io.BytesIO(header_block)
    handler.wfile=io.BytesIO()
    handler.connection=Mock()
    if stale:
        handler.headers=Message()
        handler.headers['Origin']='https://localhost:3002'
        handler._parsed_path='/powerpoint/pending'
    assert not handler.parse_request()
    raw=handler.wfile.getvalue()
    assert raw.startswith(b'HTTP/1.1 431 ')
    assert b'Access-Control-Allow-Origin' not in raw
    assert raw.count(b'HTTP/1.1 ')==1


@pytest.mark.parametrize('kind',['companion'])
@pytest.mark.parametrize('repeat',range(3))
def test_large_inflight_overload_body_preserves_503(kind,repeat):
    server = (proxy.create_server('127.0.0.1',0,max_workers=1) if kind=='proxy'
              else _BoundedThreadingHTTPServer(('127.0.0.1',0),CompanionHandler,max_handler_threads=1))
    sem = server._worker_slots if kind=='proxy' else server._handler_semaphore
    sem.acquire(False)
    thread=threading.Thread(target=server.serve_forever,daemon=True)
    thread.start()
    connection=http.client.HTTPConnection(*server.server_address,timeout=2)
    try:
        connection.request('POST','/overload-fixture',body=b'x'*(512*1024),headers={'Content-Type':'application/json'})
        response=connection.getresponse()
        assert response.status==503
        response.read()
    finally:
        connection.close()
        sem.release()
        server.shutdown()
        server.server_close()
        thread.join(2)










def test_overload_predrain_has_absolute_deadline(monkeypatch):
    from desk_plugin.companion import overload_cleanup as module
    clock=[100.0]
    monkeypatch.setattr(module.time,'monotonic',lambda:clock[0])
    owner=object.__new__(module.OverloadCleanupMixin)
    request=SlowSocket()
    def recv(size):
        clock[0]+=0.07
        request.reads+=size
        return b'x'*size
    request.recv=recv
    owner._drain_overload_request(request,100.15)
    assert request.reads==3
    assert request.timeout < 0.02
