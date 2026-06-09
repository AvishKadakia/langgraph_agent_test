#!/usr/bin/env python3
"""Host-side helper: read the web UI's Auth.js session cookie from the local
Chrome profile and print an env line ``CHAT_API_COOKIE=ets_session=<value>``.

The Docker container can't reach the host Keychain/Chrome, so `make cookie`
runs this on the host and writes the output to `.env.cookie`, which compose
loads into the container. Reads CHAT_API_URL from the environment or ./.env.

Diagnostics go to stderr; only the env line goes to stdout. Exit 1 on failure
(no cookie / wrong target / missing deps) so the Makefile can warn and continue.
"""
from __future__ import annotations

import os
import shutil
import sqlite3
import subprocess
import sys
from pathlib import Path
from urllib.parse import urlparse


def _load_env_file(path: str) -> dict[str, str]:
    out: dict[str, str] = {}
    p = Path(path)
    if not p.exists():
        return out
    for line in p.read_text().splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, v = line.split("=", 1)
            out[k.strip()] = v.strip()
    return out


def main() -> int:
    base = (os.getenv("CHAT_API_URL") or _load_env_file(".env").get("CHAT_API_URL") or "").rstrip("/")
    host = urlparse(base).hostname
    if not host:
        print("# CHAT_API_URL is not set — cannot locate a cookie", file=sys.stderr)
        return 1

    try:
        from cryptography.hazmat.backends import default_backend
        from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
        from cryptography.hazmat.primitives.hashes import SHA1
        from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    except ImportError:
        print("# host python lacks `cryptography` — paste CHAT_API_COOKIE into .env manually", file=sys.stderr)
        return 1

    db = Path.home() / "Library/Application Support/Google/Chrome/Default/Cookies"
    if not db.exists():
        print(f"# Chrome cookie store not found at {db}", file=sys.stderr)
        return 1

    pw = subprocess.run(
        ["security", "find-generic-password", "-ws", "Chrome Safe Storage"],
        capture_output=True, text=True, timeout=20,
    ).stdout.strip().encode()
    if not pw:
        print("# could not read the Chrome Safe Storage key from the Keychain", file=sys.stderr)
        return 1

    key = PBKDF2HMAC(
        algorithm=SHA1(), length=16, salt=b"saltysalt", iterations=1003, backend=default_backend()
    ).derive(pw)

    tmp = Path("/tmp/.eval_chrome_cookies.db")
    shutil.copy(db, tmp)
    con = sqlite3.connect(tmp)
    try:
        rows = con.execute(
            "SELECT name, encrypted_value FROM cookies WHERE host_key = ?", (host,)
        ).fetchall()
    finally:
        con.close()

    def decrypt(ev: bytes) -> str | None:
        if not ev:
            return None
        if ev[:3] in (b"v10", b"v11"):
            ev = ev[3:]
        dec = Cipher(algorithms.AES(key), modes.CBC(b" " * 16), backend=default_backend()).decryptor()
        pt = dec.update(ev) + dec.finalize()
        pt = pt[: -pt[-1]]  # strip PKCS#7 padding
        try:
            return pt.decode("utf-8")
        except UnicodeDecodeError:
            return pt[32:].decode("utf-8", "ignore")  # newer Chrome prepends a 32-byte domain hash

    cookies = {name: decrypt(ev) for name, ev in rows}
    session = cookies.get("ets_session")
    if not session:
        print(
            f"# no `ets_session` cookie for {host} — log in to the web UI in Chrome first "
            "(this is separate from `az login`)",
            file=sys.stderr,
        )
        return 1

    print(f"CHAT_API_COOKIE=ets_session={session}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
