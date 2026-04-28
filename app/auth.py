"""JWT auth per contract §5.

- HS256 shared secret with dispatch (`AUDIT_SERVICE_JWT_SECRET`)
- 5-minute lifetime, no refresh tokens
- Direction-aware claims: dispatch->audit issues iss=dispatch/aud=audit
  with a `file_key` claim; audit->dispatch (callback) issues iss=audit/aud=dispatch
- 30-second clock skew tolerance
"""
from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Optional

import jwt
from fastapi import Header, HTTPException, status

from .settings import settings

_ALGO = "HS256"
_MAX_SKEW_SECONDS = 30
_TOKEN_LIFETIME_SECONDS = 300  # 5 minutes


@dataclass(frozen=True)
class IncomingClaims:
    """Claims dispatch sent us when handing off a job."""

    audit_id: str
    file_key: str  # base64-encoded AES-256-GCM key for THIS audit's files


def verify_dispatch_token(authorization: str = Header(...)) -> IncomingClaims:
    """FastAPI dependency: verify a dispatch->audit JWT.

    Raises 401 on any failure. Returns the parsed claims so route handlers
    can pull `audit_id` + `file_key` directly.
    """
    if not authorization.lower().startswith("bearer "):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="missing_bearer_token",
        )

    token = authorization.split(" ", 1)[1].strip()

    try:
        payload = jwt.decode(
            token,
            settings.audit_service_jwt_secret,
            algorithms=[_ALGO],
            options={"require": ["iss", "aud", "audit_id", "iat", "exp"]},
            audience="audit",
            issuer="dispatch",
            leeway=_MAX_SKEW_SECONDS,
        )
    except jwt.ExpiredSignatureError:
        raise HTTPException(status_code=401, detail="token_expired")
    except jwt.InvalidIssuerError:
        raise HTTPException(status_code=401, detail="invalid_issuer")
    except jwt.InvalidAudienceError:
        raise HTTPException(status_code=401, detail="invalid_audience")
    except jwt.InvalidTokenError as exc:
        raise HTTPException(status_code=401, detail=f"invalid_token: {exc}")

    file_key = payload.get("file_key")
    if not file_key:
        raise HTTPException(status_code=401, detail="missing_file_key_claim")

    # Reject tokens older than our 5-min lifetime even if exp is generous
    iat = payload.get("iat", 0)
    if time.time() - iat > _TOKEN_LIFETIME_SECONDS + _MAX_SKEW_SECONDS:
        raise HTTPException(status_code=401, detail="token_too_old")

    return IncomingClaims(
        audit_id=str(payload["audit_id"]),
        file_key=str(file_key),
    )


def issue_callback_token(audit_id: str) -> str:
    """Sign an audit->dispatch token to use on the callback POST.

    No `file_key` on this direction — the callback only proves identity
    + binds the request to a specific `audit_id`.
    """
    now = int(time.time())
    payload = {
        "iss": "audit",
        "aud": "dispatch",
        "audit_id": audit_id,
        "iat": now,
        "exp": now + _TOKEN_LIFETIME_SECONDS,
    }
    return jwt.encode(payload, settings.audit_service_jwt_secret, algorithm=_ALGO)
