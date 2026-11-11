# SPDX-License-Identifier: Apache-2.0
"""Per-install localhost TLS identity. Never modifies the OS trust store."""
from __future__ import annotations

import ipaddress
import os
import shlex
import shutil
import ssl
import sys
import _ssl
import tempfile
from datetime import datetime, timedelta, timezone
from pathlib import Path


def _stdlib_ssl_context_class() -> type:
    """The standard library's SSLContext, even after ``truststore.inject_into_ssl()`` replaced ssl.SSLContext."""
    if ssl.SSLContext.__module__ == "ssl":
        return ssl.SSLContext
    for cls in _ssl._SSLContext.__subclasses__():
        if cls.__module__ == "ssl" and cls.__name__ == "SSLContext":
            return cls
    raise RuntimeError("standard library SSLContext is unavailable")


class LocalTLS:
    def __init__(self, data_dir: Path) -> None:
        self.directory = Path(data_dir) / "tls"
        self.ca_path = self.directory / "ca.pem"
        self.cert_path = self.directory / "localhost.pem"
        self.key_path = self.directory / "localhost-key.pem"
        self.ca_key_path = self.directory / "ca-key.pem"
        self._unsafe_ca_reported = False

    @staticmethod
    def _write(path: Path, content: bytes) -> None:
        # Restrict access at creation, not after writing key material.
        fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        with os.fdopen(fd, "wb") as handle:
            handle.write(content)
            handle.flush()
            os.fsync(handle.fileno())

    def ensure(self) -> ssl.SSLContext:
        # Lazy import: plugin validation/registration does not generate a CA or require a keychain.
        from cryptography import x509
        from cryptography.hazmat.primitives import hashes, serialization
        from cryptography.hazmat.primitives.asymmetric import rsa
        from cryptography.x509.oid import NameOID

        now = datetime.now(timezone.utc)
        if self.directory.is_symlink():
            raise ValueError("TLS directory must not be a symlink")
        if not self.directory.exists():
            self.directory.parent.mkdir(parents=True, exist_ok=True)
            staging = Path(tempfile.mkdtemp(prefix=".tls-", dir=self.directory.parent))
            try:
                ca_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
                leaf_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
                ca_name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "Sidecar for PowerPoint local CA " + os.urandom(8).hex())])
                ca = (x509.CertificateBuilder().subject_name(ca_name).issuer_name(ca_name)
                      .public_key(ca_key.public_key()).serial_number(x509.random_serial_number())
                      .not_valid_before(now - timedelta(minutes=5))
                      .not_valid_after(now - timedelta(minutes=5) + timedelta(days=825))
                      .add_extension(x509.BasicConstraints(ca=True, path_length=0), critical=True)
                      .add_extension(x509.KeyUsage(False, False, False, False, False, True, True, False, False), critical=True)
                      .add_extension(x509.NameConstraints([x509.DNSName("localhost"),
                          x509.IPAddress(ipaddress.ip_network("127.0.0.1/32"))], None), critical=True)
                      .add_extension(x509.SubjectKeyIdentifier.from_public_key(ca_key.public_key()), critical=False)
                      .sign(ca_key, hashes.SHA256()))
                leaf = self._leaf(ca, ca_key, leaf_key, now)
                for name, key in (("ca-key.pem", ca_key), ("localhost-key.pem", leaf_key)):
                    self._write(staging / name, key.private_bytes(serialization.Encoding.PEM,
                                serialization.PrivateFormat.PKCS8, serialization.NoEncryption()))
                self._write(staging / "ca.pem", ca.public_bytes(serialization.Encoding.PEM))
                self._write(staging / "localhost.pem", leaf.public_bytes(serialization.Encoding.PEM))
                try:
                    # Publish all four files together. A simultaneous first-run winner is reused.
                    os.rename(staging, self.directory)
                except OSError:
                    if not self.directory.is_dir():
                        raise
            finally:
                if staging.exists():
                    shutil.rmtree(staging)

        if self.directory.is_symlink():
            raise ValueError("TLS directory must not be a symlink")
        if os.name != "nt":
            os.chmod(self.directory, 0o700)
        for path in (self.ca_path, self.cert_path, self.key_path, self.ca_key_path):
            if path.is_symlink() or not path.is_file():
                raise ValueError("TLS identity is incomplete or unsafe; restore it or explicitly reset local trust")
            if os.name != "nt":
                os.chmod(path, 0o600)
        ca = x509.load_pem_x509_certificate(self.ca_path.read_bytes())
        try:
            constraints = ca.extensions.get_extension_for_class(x509.NameConstraints)
            basic = ca.extensions.get_extension_for_class(x509.BasicConstraints)
            usage = ca.extensions.get_extension_for_class(x509.KeyUsage)
            safe = (constraints.critical
                    and constraints.value == x509.NameConstraints([x509.DNSName("localhost"),
                        x509.IPAddress(ipaddress.ip_network("127.0.0.1/32"))], None)
                    and basic.critical and basic.value == x509.BasicConstraints(True, 0)
                    and usage.critical
                    and usage.value == x509.KeyUsage(False, False, False, False, False, True, True, False, False)
                    and ca.not_valid_after_utc - ca.not_valid_before_utc <= timedelta(days=825))
        except x509.ExtensionNotFound:
            safe = False
        if not safe:
            if not self._unsafe_ca_reported:
                print("UNSAFE LOCAL CA: refusing to serve; no identity has been replaced. "
                      "Stop this plugin's bridge, manually remove the old CA's trust from your login keychain "
                      "(Keychain Access) or Current User Trusted Root Certification Authorities "
                      "(Windows certificate manager), then delete only this plugin's TLS directory: "
                      f"{self.directory}. Start again and review the newly printed trust command.")
                self._unsafe_ca_reported = True
            raise ValueError("Local CA policy is unsafe; remove old trust and reset TLS directory")
        ca_key = serialization.load_pem_private_key(self.ca_key_path.read_bytes(), password=None)
        leaf_key = serialization.load_pem_private_key(self.key_path.read_bytes(), password=None)
        leaf = x509.load_pem_x509_certificate(self.cert_path.read_bytes())
        if (ca.public_key().public_bytes(serialization.Encoding.DER, serialization.PublicFormat.SubjectPublicKeyInfo)
                != ca_key.public_key().public_bytes(serialization.Encoding.DER, serialization.PublicFormat.SubjectPublicKeyInfo)
                or not ca.extensions.get_extension_for_class(x509.BasicConstraints).value.ca
                or not ca.not_valid_before_utc <= now < ca.not_valid_after_utc):
            raise ValueError("Local CA is invalid or expired; explicitly reset local trust")
        ca.verify_directly_issued_by(ca)
        leaf.verify_directly_issued_by(ca)
        if (leaf.public_key().public_bytes(serialization.Encoding.DER, serialization.PublicFormat.SubjectPublicKeyInfo)
                != leaf_key.public_key().public_bytes(serialization.Encoding.DER, serialization.PublicFormat.SubjectPublicKeyInfo)
                or now < leaf.not_valid_before_utc
                or leaf.extensions.get_extension_for_class(x509.SubjectAlternativeName).value
                .get_values_for_type(x509.DNSName) != ["localhost"]):
            raise ValueError("Localhost identity is invalid")
        if now + timedelta(days=30) >= leaf.not_valid_after_utc:
            # Renew the leaf under the same CA: no new trust step. Atomic certificate replacement.
            renewed = self._leaf(ca, ca_key, leaf_key, now)
            temp = self.directory / (".leaf-" + os.urandom(8).hex())
            try:
                self._write(temp, renewed.public_bytes(serialization.Encoding.PEM))
                os.replace(temp, self.cert_path)
            finally:
                temp.unlink(missing_ok=True)
        # Hermes injects truststore process-wide (ssl.SSLContext becomes truststore's class). That class
        # verifies the peer chain after every handshake, which fails for a server whose clients send no
        # certificate, so this server context is built from the stdlib class. The stdlib property setters
        # recurse while truststore is injected; set the floor through the C-level descriptor instead.
        context = _stdlib_ssl_context_class()(ssl.PROTOCOL_TLS_SERVER)
        _ssl._SSLContext.minimum_version.__set__(context, ssl.TLSVersion.TLSv1_2)
        context.load_cert_chain(str(self.cert_path), str(self.key_path))
        return context

    @staticmethod
    def _leaf(ca, ca_key, key, now):
        from cryptography import x509
        from cryptography.hazmat.primitives import hashes
        from cryptography.x509.oid import ExtendedKeyUsageOID, NameOID
        return (x509.CertificateBuilder()
                .subject_name(x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "localhost")]))
                .issuer_name(ca.subject).public_key(key.public_key()).serial_number(x509.random_serial_number())
                .not_valid_before(now - timedelta(minutes=5))
                .not_valid_after(min(now + timedelta(days=365), ca.not_valid_after_utc))
                .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
                .add_extension(x509.SubjectAlternativeName([x509.DNSName("localhost")]), critical=False)
                .add_extension(x509.ExtendedKeyUsage([ExtendedKeyUsageOID.SERVER_AUTH]), critical=False)
                .add_extension(x509.KeyUsage(True, False, True, False, False, False, False, False, False), critical=True)
                .add_extension(x509.AuthorityKeyIdentifier.from_issuer_public_key(ca_key.public_key()), critical=False)
                .sign(ca_key, hashes.SHA256()))

    def client_context(self) -> ssl.SSLContext:
        # Trust this install's CA only; keep hostname and certificate verification enabled.
        return ssl.create_default_context(cafile=str(self.ca_path))

    def trust_instructions(self, platform: str | None = None) -> str:
        platform = platform or sys.platform
        if platform == "win32":
            # PowerShell single-quoted literal; do not interpolate an unescaped install path.
            path = str(self.ca_path).replace("'", "''")
            command = f"certutil -user -addstore Root '{path}'"
        elif platform == "darwin":
            command = ('security add-trusted-cert -r trustRoot -k "$HOME/Library/Keychains/login.keychain-db" '
                       + shlex.quote(str(self.ca_path)))
        else:
            return f"Manually trust the local CA certificate at {self.ca_path} in your user trust store."
        return ("Manual TLS trust required once per install (never performed by the plugin). "
                "For TLS it can only vouch for localhost. The trust command does not limit the certificate to TLS, so protect the key file and remove the trust when you stop using Sidecar. "
                "Review and run this command yourself" + (" in PowerShell" if platform == "win32" else "")
                + ":\n" + command)
