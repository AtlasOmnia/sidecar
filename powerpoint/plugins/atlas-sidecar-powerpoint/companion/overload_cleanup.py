"""Fixed-capacity overload cleanup: response/read I/O never blocks acceptance.

At capacity, refusal closes immediately without HTTP bytes or additional workers.
"""
from __future__ import annotations
import queue
import socket
import threading
import time
from .request_boundaries import canonical_content_length

OVERLOAD_CLEANUP_WORKERS = 2
OVERLOAD_CLEANUP_QUEUE_CAPACITY = 8
OVERLOAD_CLEANUP_CAPACITY = OVERLOAD_CLEANUP_WORKERS + OVERLOAD_CLEANUP_QUEUE_CAPACITY
OVERLOAD_DRAIN_BYTES = 64 * 1024
OVERLOAD_DEADLINE_SECONDS = 0.2
OVERLOAD_HEADER_BYTES = 64 * 1024
OVERLOAD_BODY_BYTES = 2 * 1024 * 1024
OVERLOAD_REPLY_RESERVE_SECONDS = 0.05


class OverloadCleanupMixin:
    _overload_response = b"HTTP/1.1 503 Service Unavailable\r\nConnection: close\r\nContent-Length: 0\r\n\r\n"

    def __init__(self, *args, **kwargs):
        self._overload_cleanup_queue = queue.Queue(maxsize=OVERLOAD_CLEANUP_QUEUE_CAPACITY)
        self._overload_cleanup_lock = threading.Lock()
        self._overload_cleanup_stop = threading.Event()
        self._overload_cleanup_sockets = set()
        self._overload_cleanup_threads = []
        super().__init__(*args, **kwargs)
        try:
            for index in range(OVERLOAD_CLEANUP_WORKERS):
                thread = threading.Thread(target=self._overload_cleanup_loop,
                                          name=f"powerpoint-overload-cleanup-{index}", daemon=True)
                # Record before start so partial native starts are rolled back.
                self._overload_cleanup_threads.append(thread)
                thread.start()
        except BaseException:
            self.server_close()
            raise RuntimeError("PowerPoint transport cleanup worker startup failed") from None

    def _close_overload_request(self, request):
        try:
            self.shutdown_request(request)
        except OSError:
            self.close_request(request)

    def _queue_overload(self, request):
        admitted = False
        with self._overload_cleanup_lock:
            if (not self._overload_cleanup_stop.is_set()
                    and len(self._overload_cleanup_sockets) < OVERLOAD_CLEANUP_CAPACITY):
                self._overload_cleanup_sockets.add(request)
                try:
                    self._overload_cleanup_queue.put_nowait(request)
                    admitted = True
                except queue.Full:
                    self._overload_cleanup_sockets.discard(request)
        if not admitted:
            # Refuse before any response bytes when cleanup ownership is full.
            self._close_overload_request(request)

    def _drain_overload_request(self, request, deadline):
        head = bytearray()
        # Read only the bounded head, not speculative body/pipeline bytes.
        while not head.endswith(b"\r\n\r\n") and len(head) < OVERLOAD_HEADER_BYTES:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                return
            request.settimeout(remaining)
            chunk = request.recv(1)
            if not chunk:
                return
            head.extend(chunk)
        if not head.endswith(b"\r\n\r\n"):
            return
        lengths = []
        for line in head.split(b"\r\n")[1:]:
            key, separator, value = line.partition(b":")
            if not separator:
                continue
            if key.lower() == b"transfer-encoding":
                return
            if key.lower() == b"content-length":
                try:
                    lengths.append(value.strip().decode("ascii"))
                except UnicodeDecodeError:
                    return
        length, error = canonical_content_length(lengths, OVERLOAD_BODY_BYTES)
        if error is not None or length is None:
            return
        while length:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                return
            request.settimeout(remaining)
            chunk = request.recv(min(64 * 1024, length))
            if not chunk:
                return
            length -= len(chunk)

    def _overload_cleanup_loop(self):
        while not self._overload_cleanup_stop.is_set():
            try:
                request = self._overload_cleanup_queue.get(timeout=0.05)
            except queue.Empty:
                continue
            try:
                deadline = time.monotonic() + OVERLOAD_DEADLINE_SECONDS
                try:
                    self._drain_overload_request(request, deadline - OVERLOAD_REPLY_RESERVE_SECONDS)
                except (OSError, TimeoutError, ValueError):
                    pass
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    continue
                request.settimeout(remaining)
                request.sendall(self._overload_response)
                request.shutdown(socket.SHUT_WR)
                remaining_bytes = OVERLOAD_DRAIN_BYTES
                # Discard in-flight writes after the reply/half-close. The socket
                # is never reused; framing need not be trusted to discard bytes.
                while remaining_bytes:
                    remaining = deadline - time.monotonic()
                    if remaining <= 0:
                        break
                    request.settimeout(remaining)
                    chunk = request.recv(min(4096, remaining_bytes))
                    if not chunk:
                        break
                    remaining_bytes -= len(chunk)
            except (OSError, TimeoutError, ValueError):
                pass
            finally:
                try:
                    self._close_overload_request(request)
                finally:
                    with self._overload_cleanup_lock:
                        self._overload_cleanup_sockets.discard(request)
                    self._overload_cleanup_queue.task_done()

    def _stop_overload_cleanup(self):
        pending = []
        with self._overload_cleanup_lock:
            self._overload_cleanup_stop.set()
            while True:
                try:
                    request = self._overload_cleanup_queue.get_nowait()
                except queue.Empty:
                    break
                self._overload_cleanup_sockets.discard(request)
                pending.append(request)
            active = list(self._overload_cleanup_sockets)
        for request in pending:
            try:
                self._close_overload_request(request)
            finally:
                self._overload_cleanup_queue.task_done()
        for request in active:
            try:
                request.shutdown(socket.SHUT_RDWR)
            except OSError:
                pass
        for thread in self._overload_cleanup_threads:
            if thread.ident is not None and thread is not threading.current_thread():
                thread.join(OVERLOAD_DEADLINE_SECONDS + 0.1)

    def shutdown(self):
        super().shutdown()
        self._stop_overload_cleanup()

    def server_close(self):
        self._stop_overload_cleanup()
        super().server_close()
