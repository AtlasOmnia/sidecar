"""Word companion RPC server — ThreadingHTTPServer-based loopback companion.

Protocol v1.0 for the secure Hermes-to-Office.js pull bridge.
"""
from __future__ import annotations

import ipaddress
import logging
import socket
import ssl
import threading
import time
from http.server import ThreadingHTTPServer
from typing import Callable, Optional, cast

from .protocol import (
    INVOKE_TIMEOUT_SECONDS,
    MAX_HANDLER_THREADS,
    MAX_SESSIONS,
    WATCHDOG_HEADER_TIMEOUT_SECONDS,
)
from .session import SessionManager

_md_log = logging.getLogger("word_companion")

HandlerConcurrencySemaphore = threading.BoundedSemaphore


def _is_loopback(host: str) -> bool:
    """Return True if host is a loopback address."""
    if not host:
        return False
    try:
        addr = ipaddress.ip_address(host)
        return addr.is_loopback
    except ValueError:
        return False


class _BoundedThreadingHTTPServer(ThreadingHTTPServer):
    """ThreadingHTTPServer with bounded handler concurrency.

    Limits the number of simultaneous handler threads to prevent
    unbounded thread creation under load. Excess connections get 503.
    """

    daemon_threads = True

    def __init__(self, server_address, RequestHandlerClass,
                 max_handler_threads: int = MAX_HANDLER_THREADS,
                 request_timeout_seconds: float = 10.0,
                 bind_and_activate: bool = True, ssl_context: ssl.SSLContext | None = None):
        self._ssl_context = ssl_context
        self._clients_lock = threading.Lock()
        self._clients: set[socket.socket] = set()
        self._stopping = False
        self._handler_semaphore = HandlerConcurrencySemaphore(max_handler_threads)
        self._request_timeout_seconds = request_timeout_seconds
        super().__init__(server_address, RequestHandlerClass, bind_and_activate)

    def get_request(self):
        """Bound socket reads so slow/partial bodies cannot hold a slot forever."""
        request, client_address = super().get_request()
        request.settimeout(self._request_timeout_seconds)
        if self._ssl_context is not None:
            try:
                request = self._ssl_context.wrap_socket(request, server_side=True, do_handshake_on_connect=False)
            except OSError:
                request.close()
                raise
        with self._clients_lock:
            if self._stopping:
                request.close()
                raise OSError("Bridge stopped")
            self._clients.add(request)
        return request, client_address

    def close_request(self, request):
        with self._clients_lock:
            self._clients.discard(request)
        super().close_request(request)

    def close_clients(self):
        """Fence new accepts and interrupt TLS/keep-alive handlers from this generation."""
        with self._clients_lock:
            self._stopping = True
            clients = list(self._clients)
            self._clients.clear()
        for client in clients:
            try:
                client.shutdown(socket.SHUT_RDWR)
            except OSError:
                pass  # peer or handler already closed it
            finally:
                client.close()

    def process_request(self, request, client_address):
        """Acquire semaphore before processing; return 503 if at capacity."""
        acquired = self._handler_semaphore.acquire(blocking=False)
        if not acquired:
            request_socket = cast(socket.socket, request)
            # Windows sends a reset when a socket is closed with unread request
            # bytes. Send and half-close first, then drain within absolute
            # time/byte limits so a trickling peer cannot hold the accept loop.
            try:
                timeout = min(self._request_timeout_seconds, 0.2)
                request_socket.settimeout(timeout)
                body = b'{"ok":false,"error":"Server busy"}'
                headers = (
                    "HTTP/1.1 503 Service Unavailable\r\n"
                    "Content-Type: application/json\r\n"
                    "Connection: close\r\n"
                    f"Content-Length: {len(body)}\r\n"
                    "\r\n"
                ).encode("ascii")
                request_socket.sendall(headers + body)
                try:
                    request_socket.shutdown(socket.SHUT_WR)
                except OSError:
                    pass
                deadline = time.monotonic() + timeout
                remaining_bytes = 64 * 1024
                while remaining_bytes:
                    remaining_seconds = deadline - time.monotonic()
                    if remaining_seconds <= 0:
                        break
                    request_socket.settimeout(remaining_seconds)
                    chunk = request_socket.recv(min(4096, remaining_bytes))
                    if not chunk:
                        break
                    remaining_bytes -= len(chunk)
            except (OSError, TimeoutError):
                pass
            finally:
                self.close_request(request)
            return

        try:
            super().process_request(request, client_address)
        except Exception:
            self._handler_semaphore.release()
            raise

    def process_request_thread(self, request, client_address):
        """Release the handler slot only after the spawned thread finishes."""
        try:
            if self._ssl_context is not None:
                try:
                    cast(ssl.SSLSocket, request).do_handshake()
                except OSError as exc:
                    # Metadata only: never certificate or key material.
                    _md_log.info("method=UNKNOWN route=TLS_HANDSHAKE error=%s", type(exc).__name__)
                    self.close_request(request)
                    return
            super().process_request_thread(request, client_address)
        finally:
            self._handler_semaphore.release()


class WordCompanionServer:
    """Loopback-only ThreadingHTTPServer for the Word companion RPC protocol."""

    DEFAULT_HOST = "127.0.0.1"
    DEFAULT_PORT = 8651

    def __init__(self, host: Optional[str] = None, port: Optional[int] = None,
                 plugin_key: Optional[str] = None,
                 max_sessions: int = MAX_SESSIONS,
                 max_handler_threads: int = MAX_HANDLER_THREADS,
                 clock: Optional[Callable[[], float]] = None,
                 invoke_timeout_seconds: float = INVOKE_TIMEOUT_SECONDS,
                 request_timeout_seconds: float = WATCHDOG_HEADER_TIMEOUT_SECONDS,
                 ssl_context: ssl.SSLContext | None = None) -> None:
        self._ssl_context = ssl_context
        if host is None:
            host = self.DEFAULT_HOST
        elif not host:
            raise ValueError("Companion host must not be empty")

        port = port if port is not None else self.DEFAULT_PORT

        if not _is_loopback(host):
            raise ValueError(f"Companion must bind to a loopback address, got {host!r}")

        self._host = host
        self._port = port
        self.plugin_key = plugin_key
        self._clock = clock or time.monotonic
        self.sessions = SessionManager(clock=self._clock, max_sessions=max_sessions)
        self._max_handler_threads = max_handler_threads
        if invoke_timeout_seconds <= 0:
            raise ValueError("invoke_timeout_seconds must be positive")
        self.invoke_timeout_seconds = float(invoke_timeout_seconds)
        if request_timeout_seconds <= 0:
            raise ValueError("request_timeout_seconds must be positive")
        self._request_timeout_seconds = float(request_timeout_seconds)
        self._server: Optional[_BoundedThreadingHTTPServer] = None
        self._thread: Optional[threading.Thread] = None
        self._started_at: float = 0.0

        # Configure metadata-only logger
        self._setup_logging()

    def _setup_logging(self) -> None:
        # Runs inside the Hermes process: never attach a handler or change levels. Records stay
        # metadata-only (route template, status, elapsed) and go wherever Hermes sends logs.
        logging.getLogger("word_companion")

    def _handler_base(self) -> type:
        """Handler class the server binds; the plugin's desk bridge overrides this seam."""
        from .handler import CompanionHandler
        return CompanionHandler

    def start(self) -> None:
        """Start the server in a background daemon thread."""
        handler = self._handler_base().with_server(self)
        self._server = _BoundedThreadingHTTPServer(
            (self._host, self._port), handler,
            max_handler_threads=self._max_handler_threads,
            request_timeout_seconds=self._request_timeout_seconds,
            ssl_context=self._ssl_context,
        )
        self._started_at = self._clock()
        self._thread = threading.Thread(target=self._server.serve_forever, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        """Stop the server, cancel all tickets, and wait for the thread to join."""
        # Cancel all outstanding tickets
        self.sessions.shutdown()

        if self._server:
            self._server.close_clients()
            self._server.shutdown()
            self._server.server_close()
        if self._thread:
            self._thread.join(timeout=5)

    def server_address(self) -> tuple[str, int]:
        """Return the actual (host, port) the server is listening on."""
        if self._server is None:
            return (self._host, self._port)
        return self._server.server_address[:2]  # type: ignore[return-value]

    @property
    def request_timeout_seconds(self) -> float:
        """Absolute header/body read deadline for each parsed request."""
        return self._request_timeout_seconds

    @property
    def uptime(self) -> float:
        """Server uptime in seconds (non-negative elapsed monotonic)."""
        if self._started_at == 0.0:
            return 0.0
        elapsed = self._clock() - self._started_at
        return max(0.0, elapsed)

    def now(self) -> float:
        """Return the server's injected monotonic clock value."""
        return self._clock()
