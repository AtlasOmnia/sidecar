"""One absolute downstream write budget for status, headers, body and flush."""
from __future__ import annotations
import socket
import time
import json

RESPONSE_WRITE_DEADLINE_SECONDS = 1.0


class ResponseWriteAborted(OSError):
    pass


class _DeadlineWriter:
    def __init__(self, owner, stream):
        self.owner, self.stream = owner, stream
    def write(self, data):
        return self.owner._write_response(data)
    def flush(self):
        return self.owner._flush_response()
    def __getattr__(self, name):
        return getattr(self.stream, name)


class ResponseDeadlineMixin:
    def setup(self):
        super().setup()
        self._response_deadline = None
        self.wfile = _DeadlineWriter(self, self.wfile)

    def _begin_response_write(self):
        if getattr(self, '_response_deadline', None) is None:
            self._response_deadline = time.monotonic() + getattr(self, '_response_write_budget', RESPONSE_WRITE_DEADLINE_SECONDS)

    def _prepare_response_write(self):
        self._begin_response_write()
        remaining = self._response_deadline - time.monotonic()
        if remaining <= 0:
            raise ResponseWriteAborted('response deadline exceeded')
        try:
            self.connection.settimeout(min(remaining, RESPONSE_WRITE_DEADLINE_SECONDS))
        except (OSError, ValueError) as exc:
            raise ResponseWriteAborted('response timeout failed') from exc

    def _stream(self):
        return self.wfile.stream if isinstance(self.wfile, _DeadlineWriter) else self.wfile

    def _write_response(self, data):
        self._prepare_response_write()
        try:
            return self._stream().write(data)
        except (OSError, ValueError) as exc:
            raise ResponseWriteAborted('response write failed') from exc

    def _flush_response(self):
        self._prepare_response_write()
        try:
            return self._stream().flush()
        except (OSError, ValueError) as exc:
            raise ResponseWriteAborted('response flush failed') from exc

    def _abort_response_write(self):
        self.close_connection = True
        self._response_deadline = None
        connection = getattr(self, 'connection', None)
        if connection is not None:
            try:
                connection.shutdown(socket.SHUT_RDWR)
            except (OSError, ValueError):
                pass
            try:
                connection.close()
            except (OSError, ValueError):
                pass

    def _finish_response_write(self):
        self._response_deadline = None
        if not getattr(self, 'close_connection', False):
            srv = getattr(self, '_server_ref', None)
            try:
                self.connection.settimeout(srv.request_timeout_seconds if srv else 10.0)
            except (OSError, ValueError):
                self._abort_response_write()

    def flush_headers(self):
        if hasattr(self, '_headers_buffer'):
            self._write_response(b''.join(self._headers_buffer))
            self._headers_buffer = []
            if getattr(self, '_response_is_final', True):
                self._response_committed = True

    def send_response_only(self, code, message=None):
        self._begin_response_write()
        self._response_is_final = code >= 200
        super().send_response_only(code, message)

    def send_error(self, code, message=None, explain=None):
        # Do not echo a request line/method from the base parser's HTML errors.
        self.close_connection = True
        # Parser errors can precede self.headers and must never consult CORS or
        # route-auth state from either this incomplete or a previous request.
        body = json.dumps({'ok': False,
                           'error': self.responses.get(code, ('Internal Server Error',''))[0],
                           'category': self._error_category(code)}).encode()
        self._emit_response(code, body, time.monotonic(),
                            extra_headers={'Connection':'close'}, preflight_cors=False)

    def finish(self):
        try:
            super().finish()
        except ResponseWriteAborted:
            self._abort_response_write()
