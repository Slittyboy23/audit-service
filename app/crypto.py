"""AES-256-GCM en/decryption per contract §6.

Encrypted blob format (matches dispatch's existing convention):
  [12-byte IV] [ciphertext] [16-byte GCM tag]

The per-audit AES key arrives in the JWT `file_key` claim, base64-encoded.
Decrypt happens in-memory only — keys are never written to disk.
"""
from __future__ import annotations

import base64
import os
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

_IV_LEN = 12  # GCM standard


def _decode_key(b64_key: str) -> bytes:
    """Decode the JWT-delivered key. Accept both URL-safe and standard base64."""
    # Add padding if missing (urlsafe encoders sometimes strip it)
    padding = "=" * (-len(b64_key) % 4)
    try:
        return base64.urlsafe_b64decode(b64_key + padding)
    except Exception:
        return base64.b64decode(b64_key + padding)


def decrypt(blob: bytes, b64_key: str) -> bytes:
    """Decrypt an encrypted blob produced by dispatch."""
    key = _decode_key(b64_key)
    if len(key) != 32:
        raise ValueError(f"file_key must decode to 32 bytes, got {len(key)}")
    if len(blob) < _IV_LEN + 16:
        raise ValueError("encrypted blob too short")

    iv, ciphertext = blob[:_IV_LEN], blob[_IV_LEN:]
    return AESGCM(key).decrypt(iv, ciphertext, associated_data=None)


def encrypt(plaintext: bytes, b64_key: str) -> bytes:
    """Encrypt the output workbook with the same per-audit key for the callback."""
    key = _decode_key(b64_key)
    if len(key) != 32:
        raise ValueError(f"file_key must decode to 32 bytes, got {len(key)}")

    iv = os.urandom(_IV_LEN)
    ciphertext = AESGCM(key).encrypt(iv, plaintext, associated_data=None)
    return iv + ciphertext
