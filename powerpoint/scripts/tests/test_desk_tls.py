"""D4: real TLS, trust confined to a disposable test client, never an OS keychain."""
import http.client
import os
import socket
import ssl
import time
from pathlib import Path
from datetime import datetime, timedelta, timezone
from concurrent.futures import ThreadPoolExecutor

import pytest
from cryptography import x509
from cryptography.hazmat.primitives import serialization

from test_desk_bridge_security import APP, ORIGIN, make_runtime, running
from desk_plugin.desk_bridge import DeskError
from desk_plugin.desk_tls import LocalTLS


def tls_request(rt, host=None, origin=ORIGIN, context=None):
    context = context or ssl.create_default_context(cafile=str(rt.store.data_dir / "tls" / "ca.pem"))
    conn = http.client.HTTPSConnection("localhost", rt.bridge.port, context=context, timeout=3)
    try:
        conn.request("GET", f"/{APP}/hello", headers={
            "Host": host or f"localhost:{rt.bridge.port}", "Origin": origin})
        response = conn.getresponse()
        return response.status, dict(response.getheaders()), response.read()
    finally:
        conn.close()


def test_d4_tls_first_run_and_persistent_identity(tmp_path, caplog, capsys):
    rt = make_runtime(tmp_path)
    assert not (tmp_path / "tls").exists()  # construction/validation does not generate keys
    with running(rt):
        assert rt.bridge.base_url == f"https://localhost:{rt.bridge.port}"
        assert rt.bridge._server.server_address()[0] == "127.0.0.1"
        assert tls_request(rt)[0] == 200
        ca = (tmp_path / "tls" / "ca.pem").read_bytes()
        leaf = x509.load_pem_x509_certificate((tmp_path / "tls" / "localhost.pem").read_bytes())
        assert leaf.extensions.get_extension_for_class(x509.SubjectAlternativeName).value.get_values_for_type(x509.DNSName) == ["localhost"]
        for name in ("ca-key.pem", "localhost-key.pem"):
            path = tmp_path / "tls" / name
            serialization.load_pem_private_key(path.read_bytes(), password=None)
            if os.name != "nt":
                assert path.stat().st_mode & 0o777 == 0o600
        assert rt.ensure_started() == rt.bridge.port
    with running(make_runtime(tmp_path)) as restarted:
        assert (tmp_path / "tls" / "ca.pem").read_bytes() == ca
        assert tls_request(restarted)[0] == 200
    logged = caplog.text + capsys.readouterr().out
    assert "PRIVATE KEY" not in logged
    assert "add-trusted-cert" in logged or "certutil" in logged


def test_d4_trust_required_plaintext_refused_and_strict_gates(tmp_path):
    with running(make_runtime(tmp_path)) as rt:
        with pytest.raises(ssl.SSLCertVerificationError):
            tls_request(rt, context=ssl.create_default_context())
        for host in (f"127.0.0.1:{rt.bridge.port}", "localhost", f"localhost:{rt.bridge.port + 1}", "evil.example"):
            assert tls_request(rt, host=host)[0] == 421
        status, headers, _ = tls_request(rt, origin="https://evil.example")
        assert status == 403
        assert "Access-Control-Allow-Origin" not in headers
        with socket.create_connection(("127.0.0.1", rt.bridge.port), timeout=3) as conn:
            conn.sendall(b"GET /powerpoint/hello HTTP/1.1\r\nHost: localhost\r\n\r\n")
            try:
                result = conn.recv(1024)
            except ConnectionResetError:
                result = b""
            assert b"HTTP/" not in result
        assert tls_request(rt)[0] == 200


def test_d4_independent_installs_have_different_ca(tmp_path):
    with running(make_runtime(tmp_path / "a")) as a, running(make_runtime(tmp_path / "b")) as b:
        assert (tmp_path / "a/tls/ca.pem").read_bytes() != (tmp_path / "b/tls/ca.pem").read_bytes()
        with pytest.raises(ssl.SSLCertVerificationError):
            tls_request(b, context=a.bridge.tls.client_context())


@pytest.mark.parametrize("platform,needle", [("darwin", "security add-trusted-cert"), ("win32", "certutil -user -addstore Root")])
def test_d4_manual_trust_command_user_scope_and_quoted_path(tmp_path, platform, needle):
    tls = LocalTLS(tmp_path / "user's data with spaces")
    text = tls.trust_instructions(platform)
    assert text.count(needle) == 1
    assert "sudo" not in text and "System.keychain" not in text
    assert "ca-key.pem" not in text and "localhost-key.pem" not in text
    assert "never performed" in text
    if platform == "darwin":
        import shlex
        assert shlex.split(text.splitlines()[-1])[-1] == str(tls.ca_path)
        assert "login.keychain-db" in text
    else:
        assert "user''s data with spaces" in text
    assert not tls.directory.exists()


def test_d4_incomplete_identity_fails_closed_without_replacing_ca(tmp_path):
    tls = LocalTLS(tmp_path)
    tls.ensure()
    ca = tls.ca_path.read_bytes()
    tls.key_path.unlink()
    rt = make_runtime(tmp_path)
    with pytest.raises(DeskError, match="Local TLS identity") as exc:
        rt.ensure_started()
    assert exc.value.category == "TLS"
    assert rt.bridge.port is None
    assert tls.ca_path.read_bytes() == ca
    assert not tls.key_path.exists()


def test_d4_leaf_renewal_preserves_ca_and_key(tmp_path):
    tls = LocalTLS(tmp_path)
    tls.ensure()
    ca_bytes = tls.ca_path.read_bytes()
    key_bytes = tls.key_path.read_bytes()
    ca = x509.load_pem_x509_certificate(ca_bytes)
    ca_key = serialization.load_pem_private_key(tls.ca_key_path.read_bytes(), password=None)
    key = serialization.load_pem_private_key(key_bytes, password=None)
    old = tls._leaf(ca, ca_key, key, datetime.now(timezone.utc) - timedelta(days=360))
    tls.cert_path.write_bytes(old.public_bytes(serialization.Encoding.PEM))
    tls.ensure()
    assert tls.ca_path.read_bytes() == ca_bytes
    assert tls.key_path.read_bytes() == key_bytes
    renewed = x509.load_pem_x509_certificate(tls.cert_path.read_bytes())
    assert renewed.not_valid_after_utc > old.not_valid_after_utc
    with running(make_runtime(tmp_path)) as rt:
        assert tls_request(rt)[0] == 200


def test_d5_ca_constraints_and_validity(tmp_path):
    import ipaddress
    from cryptography.x509.oid import ExtendedKeyUsageOID
    tls = LocalTLS(tmp_path)
    tls.ensure()
    ca = x509.load_pem_x509_certificate(tls.ca_path.read_bytes())
    extension = ca.extensions.get_extension_for_class(x509.NameConstraints)
    assert extension.critical
    assert extension.value.permitted_subtrees == [
        x509.DNSName("localhost"), x509.IPAddress(ipaddress.ip_network("127.0.0.1/32"))]
    assert extension.value.excluded_subtrees is None
    assert ca.not_valid_after_utc - ca.not_valid_before_utc <= timedelta(days=825)
    basic = ca.extensions.get_extension_for_class(x509.BasicConstraints)
    assert basic.critical and basic.value == x509.BasicConstraints(True, 0)
    usage = ca.extensions.get_extension_for_class(x509.KeyUsage)
    assert usage.critical
    assert usage.value == x509.KeyUsage(False, False, False, False, False, True, True, False, False)
    leaf = x509.load_pem_x509_certificate(tls.cert_path.read_bytes())
    assert leaf.not_valid_after_utc - leaf.not_valid_before_utc <= timedelta(days=397)
    assert leaf.extensions.get_extension_for_class(x509.BasicConstraints).value.ca is False
    assert list(leaf.extensions.get_extension_for_class(x509.ExtendedKeyUsage).value) == [ExtendedKeyUsageOID.SERVER_AUTH]
    assert list(leaf.extensions.get_extension_for_class(x509.SubjectAlternativeName).value) == [x509.DNSName("localhost")]


@pytest.mark.parametrize("policy", ["missing", "noncritical", "broad", "extra", "excluded", "long"])
def test_d5_legacy_or_incorrect_ca_refuses_without_rotation(tmp_path, capsys, policy):
    import ipaddress
    from cryptography.hazmat.primitives import hashes
    tls = LocalTLS(tmp_path)
    tls.ensure()
    original = x509.load_pem_x509_certificate(tls.ca_path.read_bytes())
    key = serialization.load_pem_private_key(tls.ca_key_path.read_bytes(), password=None)
    from cryptography.hazmat.primitives.asymmetric import rsa
    assert isinstance(key, rsa.RSAPrivateKey)
    builder = (x509.CertificateBuilder().subject_name(original.subject).issuer_name(original.issuer)
               .public_key(key.public_key()).serial_number(original.serial_number)
               .not_valid_before(original.not_valid_before_utc)
               .not_valid_after(original.not_valid_before_utc + timedelta(days=826) if policy == "long"
                                else original.not_valid_after_utc))
    for extension in original.extensions:
        if not isinstance(extension.value, x509.NameConstraints):
            builder = builder.add_extension(extension.value, extension.critical)
    if policy != "missing":
        permitted = [x509.DNSName("localhost"), x509.IPAddress(ipaddress.ip_network("127.0.0.1/32"))]
        if policy == "broad":
            permitted[1] = x509.IPAddress(ipaddress.ip_network("127.0.0.0/8"))
        if policy == "extra":
            permitted.append(x509.DNSName("example.com"))
        builder = builder.add_extension(x509.NameConstraints(permitted,
            [x509.DNSName("blocked.localhost")] if policy == "excluded" else None), policy != "noncritical")
    legacy = builder.sign(key, hashes.SHA256()).public_bytes(serialization.Encoding.PEM)
    tls.ca_path.write_bytes(legacy)
    rt = make_runtime(tmp_path)
    for _ in range(2):
        with pytest.raises(DeskError, match="Local TLS identity"):
            rt.ensure_started()
        assert rt.bridge.port is None
        assert tls.ca_path.read_bytes() == legacy
    output = capsys.readouterr().out
    assert output.count("UNSAFE LOCAL CA") == 1
    assert "remove" in output and "trust" in output and "delete" in output
    assert str(tls.directory) in output
    assert "PRIVATE KEY" not in output


@pytest.mark.parametrize("name", ["localhost", "example.com", "127.0.0.2", "::1"])
def test_d5_ca_enforces_name_constraints_in_verifier(tmp_path, name):
    import ipaddress
    from cryptography.hazmat.primitives import hashes
    from cryptography.hazmat.primitives.asymmetric import rsa
    from cryptography.x509.oid import ExtendedKeyUsageOID, NameOID
    from cryptography.x509.verification import PolicyBuilder, Store, VerificationError
    tls = LocalTLS(tmp_path)
    tls.ensure()
    ca = x509.load_pem_x509_certificate(tls.ca_path.read_bytes())
    key = serialization.load_pem_private_key(tls.ca_key_path.read_bytes(), password=None)
    assert isinstance(key, rsa.RSAPrivateKey)
    now = datetime.now(timezone.utc)
    san = x509.IPAddress(ipaddress.ip_address(name)) if ":" in name or name[0].isdigit() else x509.DNSName(name)
    # Deliberately exercise a CA-key holder minting an unrelated leaf, not just a hostname mismatch.
    cert = (x509.CertificateBuilder().subject_name(x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, name)]))
            .issuer_name(ca.subject).public_key(key.public_key()).serial_number(x509.random_serial_number())
            .not_valid_before(now - timedelta(minutes=5)).not_valid_after(now + timedelta(days=1))
            .add_extension(x509.BasicConstraints(False, None), critical=True)
            .add_extension(x509.SubjectAlternativeName([san]), critical=False)
            .add_extension(x509.ExtendedKeyUsage([ExtendedKeyUsageOID.SERVER_AUTH]), critical=False)
            .add_extension(x509.AuthorityKeyIdentifier.from_issuer_public_key(key.public_key()), critical=False)
            .sign(key, hashes.SHA256()))
    verifier = PolicyBuilder().store(Store([ca])).build_server_verifier(san)
    if name == "localhost":
        assert verifier.verify(cert, [])
    else:
        with pytest.raises(VerificationError, match="name constraint"):
            verifier.verify(cert, [])


def test_d4_parallel_first_run_reuses_one_complete_identity(tmp_path):
    def create(_):
        tls = LocalTLS(tmp_path)
        tls.ensure()
        return tls.ca_path.read_bytes()
    with ThreadPoolExecutor(max_workers=3) as pool:
        identities = list(pool.map(create, range(3)))
    assert len(set(identities)) == 1
    assert not list(tmp_path.glob(".tls-*"))


def test_d4_duplicate_host_and_stalled_handshake(tmp_path):
    with running(make_runtime(tmp_path)) as rt:
        with socket.create_connection(("127.0.0.1", rt.bridge.port), timeout=3) as stalled:
            # Send no TLS handshake. Other clients must still be accepted promptly.
            assert tls_request(rt)[0] == 200
        with rt.bridge.tls.client_context().wrap_socket(
                socket.create_connection(("127.0.0.1", rt.bridge.port), timeout=3), server_hostname="localhost") as conn:
            conn.sendall((f"GET /{APP}/hello HTTP/1.1\r\nHost: localhost:{rt.bridge.port}\r\n"
                          f"Host: localhost:{rt.bridge.port}\r\nOrigin: {ORIGIN}\r\nConnection: close\r\n\r\n").encode())
            assert b" 421 " in conn.recv(4096)


@pytest.fixture
def truststore_injected():
    """Hermes injects truststore process-wide (agent/ssl_verify.py); the bridge must still serve."""
    truststore = pytest.importorskip("truststore")
    stdlib_context = ssl.SSLContext
    truststore.inject_into_ssl()
    try:
        assert ssl.SSLContext is not stdlib_context
        yield stdlib_context
    finally:
        truststore.extract_from_ssl()


def test_s5_bridge_serves_with_truststore_injected(tmp_path, truststore_injected):
    with running(make_runtime(tmp_path)) as rt:
        client = truststore_injected(ssl.PROTOCOL_TLS_CLIENT)
        client.load_verify_locations(cafile=str(tmp_path / "tls" / "ca.pem"))
        status, headers, _ = tls_request(rt, context=client)
        assert status == 200
        assert headers["Access-Control-Allow-Origin"] == ORIGIN
        assert client.minimum_version >= ssl.TLSVersion.TLSv1_2
        legacy = truststore_injected(ssl.PROTOCOL_TLS_CLIENT)
        legacy.load_verify_locations(cafile=str(tmp_path / "tls" / "ca.pem"))
        import _ssl
        _ssl._SSLContext.minimum_version.__set__(legacy, ssl.TLSVersion.TLSv1)
        _ssl._SSLContext.maximum_version.__set__(legacy, ssl.TLSVersion.TLSv1_1)
        with pytest.raises((ssl.SSLError, OSError)):
            tls_request(rt, context=legacy)


def test_s5_failed_handshake_is_logged_not_silent(tmp_path, caplog):
    caplog.set_level("INFO", logger="powerpoint_companion")
    with running(make_runtime(tmp_path)) as rt:
        with pytest.raises(ssl.SSLCertVerificationError):
            tls_request(rt, context=ssl.create_default_context())
        deadline = time.monotonic() + 3
        while "route=TLS_HANDSHAKE" not in caplog.text and time.monotonic() < deadline:
            time.sleep(0.05)
    assert "route=TLS_HANDSHAKE error=" in caplog.text
    assert "PRIVATE KEY" not in caplog.text


@pytest.mark.parametrize("route,method", [("heartbeat", "POST"), ("pending", "GET"), ("result", "POST"),
                                           ("session", "POST")])
def test_s5_session_route_preflight_allows_the_configured_origin(tmp_path, route, method):
    """A sidebar on the configured add-in origin must pass the preflight of every pane route, not only
    the connect routes; the session headers make each of these requests preflighted."""
    from desk_plugin.companion.protocol import OPTIONS_ROUTES
    assert OPTIONS_ROUTES == {f"/{APP}/{name}" for name in ("session", "heartbeat", "pending", "result")}
    with running(make_runtime(tmp_path)) as rt:
        def preflight(origin):
            conn = http.client.HTTPSConnection("localhost", rt.bridge.port, timeout=3,
                                               context=rt.bridge.tls.client_context())
            try:
                conn.request("OPTIONS", f"/{APP}/{route}", headers={
                    "Host": f"localhost:{rt.bridge.port}", "Origin": origin,
                    "Access-Control-Request-Method": method,
                    "Access-Control-Request-Headers": "content-type,x-hermes-session-key,x-hermes-session-id,x-hermes-pane-id"})
                response = conn.getresponse()
                response.read()
                return response.status, {k.lower(): v for k, v in response.getheaders()}
            finally:
                conn.close()
        status, headers = preflight(ORIGIN)
        assert status == 204
        assert headers.get("access-control-allow-origin") == ORIGIN
        allowed = headers.get("access-control-allow-headers", "").lower()
        for name in ("content-type", "x-hermes-session-key", "x-hermes-session-id", "x-hermes-pane-id"):
            assert name in allowed
        status, headers = preflight("https://evil.example")
        assert "access-control-allow-origin" not in headers
