"""FastAPI app implementing the 3 endpoints from contract §4.

  POST /v1/audits         — dispatch hands off a job, we 202 + process async
  POST /v1/audits/callback — (NOT here; that's dispatch's endpoint we POST TO)
  GET  /v1/audits/{id}    — polling fallback (contract §2 decision 5)
  GET  /v1/health         — liveness probe (no auth, contract §4.3)

The service is stateless — no DB, no persistent volume. In-memory job
tracking exists ONLY to power the polling fallback. If the process restarts,
in-flight jobs are lost; dispatch's 15-min callback timeout (§8) catches it.
"""
from __future__ import annotations

import asyncio
import logging
import time
from datetime import datetime, timezone
from typing import Dict, Optional

import httpx
from fastapi import (
    BackgroundTasks,
    Depends,
    FastAPI,
    File,
    Form,
    HTTPException,
    Request,
    UploadFile,
)
from fastapi.responses import JSONResponse

from .audit_engine import AuditEngineError, AuditInputs, run_audit
from .auth import IncomingClaims, issue_callback_token, verify_dispatch_token
from .crypto import decrypt, encrypt
from .models import AcceptedResponse, AuditStatusResponse, HealthResponse
from .settings import settings

# ------------------------------------------------------------------
# Setup
# ------------------------------------------------------------------
logging.basicConfig(
    level=settings.log_level,
    format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
)
log = logging.getLogger("audit-service")

app = FastAPI(
    title="Apollo Audit Service",
    version=settings.service_version,
    docs_url=None,  # docs disabled — server-to-server only, no humans hitting this
    redoc_url=None,
    openapi_url=None,
)

_started_at = time.time()

# In-memory status table for polling fallback (contract §4 / decision 5).
# Lost on restart by design — dispatch's callback-timeout catches orphans.
_jobs: Dict[str, AuditStatusResponse] = {}
_jobs_lock = asyncio.Lock()
_active_count = 0
_active_lock = asyncio.Lock()


# ------------------------------------------------------------------
# Endpoints
# ------------------------------------------------------------------
@app.get("/v1/health", response_model=HealthResponse)
async def health() -> HealthResponse:
    """Public liveness probe. No auth, no sensitive info."""
    return HealthResponse(
        status="ok",
        version=settings.service_version,
        uptime_s=int(time.time() - _started_at),
    )


@app.get("/v1/audits/{audit_id}", response_model=AuditStatusResponse)
async def get_audit_status(audit_id: str) -> AuditStatusResponse:
    """Polling fallback — dispatch only hits this if the callback never arrives."""
    async with _jobs_lock:
        job = _jobs.get(audit_id)
    if not job:
        # Stateless service: if we don't know it, we never accepted it (or restarted)
        raise HTTPException(status_code=404, detail="audit_not_found")
    return job


@app.post("/v1/preview")
async def preview_file(
    kind: str = Form(...),                 # 'rent_roll' | 'vehicles' | 'myvip'
    filename: str = Form(...),
    file: UploadFile = File(...),
    column_mappings: Optional[str] = Form(None),
    claims: IncomingClaims = Depends(verify_dispatch_token),
) -> dict:
    """Parse a single file and return summary counts for the wizard preview UI.

    Synchronous (blocks for parse duration; ~1-3s for typical files). No workbook
    generation, no callback. The same per-audit JWT auth/encryption applies.

    Response shape varies by kind:
      rent_roll → { total_units, occupied_units, vacant_units, model_units,
                    ntv_units, future_residents }
      vehicles  → { total_vehicles, total_units }
      myvip     → { total_units }
    """
    if kind not in {"rent_roll", "vehicles", "myvip"}:
        raise HTTPException(status_code=400, detail=f"unknown_preview_kind: {kind}")

    raw = await file.read()
    try:
        plain = decrypt(raw, claims.file_key)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"decryption_failed: {exc}")

    from .parsers import parse_rent_roll, parse_vehicle_data
    from collections import Counter

    try:
        if kind == "rent_roll":
            entries = parse_rent_roll(plain, filename)
            primary = {}
            future = 0
            for e in entries:
                if e.is_future:
                    future += 1
                else:
                    primary[e.apt] = e
            counts = Counter(e.status for e in primary.values())
            return {
                "total_units": len(primary),
                "occupied_units": counts.get("occupied", 0),
                "vacant_units": counts.get("vacant", 0),
                "model_units": counts.get("model", 0),
                "ntv_units": counts.get("ntv", 0),
                "future_residents": future,
            }
        else:
            vehicles = parse_vehicle_data(plain, filename)
            return {
                "total_vehicles": len(vehicles),
                "total_units": len({v.apt for v in vehicles}),
            }
    except AuditEngineError as e:
        raise HTTPException(status_code=400, detail={"code": e.code, "message": e.message})


@app.post("/v1/audits", status_code=202, response_model=AcceptedResponse)
async def submit_audit(
    background: BackgroundTasks,
    request: Request,
    audit_id: str = Form(...),
    property_name: str = Form(...),
    property_uuid: str = Form(...),
    callback_url: str = Form(...),
    rent_roll: UploadFile = File(...),
    rent_roll_filename: str = Form(...),
    vehicle_data: UploadFile = File(...),
    vehicle_data_filename: str = Form(...),
    myvip_data: Optional[UploadFile] = File(None),
    myvip_data_filename: Optional[str] = Form(None),
    column_mappings: Optional[str] = Form(None),
    claims: IncomingClaims = Depends(verify_dispatch_token),
) -> AcceptedResponse:
    """Accept a job, return 202 immediately, process in background."""

    # The audit_id in the JWT MUST match the form body — defends against a
    # token issued for one audit being replayed against another's payload.
    if claims.audit_id != audit_id:
        raise HTTPException(status_code=401, detail="audit_id_mismatch")

    # Reject only on duplicate IN-FLIGHT or completed jobs we still remember.
    async with _jobs_lock:
        if audit_id in _jobs:
            raise HTTPException(status_code=409, detail="duplicate_audit_id")

    # Soft cap on concurrency (contract §9). Hard rate-limit lives at the LB.
    async with _active_lock:
        if _active_count >= settings.max_concurrent_audits:
            raise HTTPException(status_code=503, detail="queue_full")

    # Eagerly read uploads into memory — files are small (≤50 MB total per §4.1)
    # and the service is stateless, so we don't spool to disk.
    rent_roll_bytes = await rent_roll.read()
    vehicle_data_bytes = await vehicle_data.read()
    myvip_bytes = await myvip_data.read() if myvip_data is not None else None

    accepted_at = datetime.now(timezone.utc)
    async with _jobs_lock:
        _jobs[audit_id] = AuditStatusResponse(
            audit_id=audit_id,
            status="processing",
            accepted_at=accepted_at,
        )

    background.add_task(
        _process_audit,
        audit_id=audit_id,
        property_name=property_name,
        property_uuid=property_uuid,
        callback_url=callback_url,
        file_key=claims.file_key,
        rent_roll_bytes=rent_roll_bytes,
        rent_roll_filename=rent_roll_filename,
        vehicle_data_bytes=vehicle_data_bytes,
        vehicle_data_filename=vehicle_data_filename,
        myvip_bytes=myvip_bytes,
        myvip_filename=myvip_data_filename,
        column_mappings_raw=column_mappings,
    )

    return AcceptedResponse(audit_id=audit_id, accepted_at=accepted_at)


# ------------------------------------------------------------------
# Background processing
# ------------------------------------------------------------------
async def _process_audit(
    *,
    audit_id: str,
    property_name: str,
    property_uuid: str,
    callback_url: str,
    file_key: str,
    rent_roll_bytes: bytes,
    rent_roll_filename: str,
    vehicle_data_bytes: bytes,
    vehicle_data_filename: str,
    myvip_bytes: Optional[bytes],
    myvip_filename: Optional[str],
    column_mappings_raw: Optional[str],
) -> None:
    """Run the audit engine, encrypt the result, POST it back to dispatch."""
    global _active_count
    async with _active_lock:
        _active_count += 1

    started = time.monotonic()
    error_code: Optional[str] = None
    error_message: Optional[str] = None
    workbook_enc: Optional[bytes] = None
    workbook_filename: Optional[str] = None
    metrics: Optional[dict] = None

    try:
        # 1) Decrypt inputs in-memory
        rent_roll_plain = decrypt(rent_roll_bytes, file_key)
        vehicle_plain = decrypt(vehicle_data_bytes, file_key)
        myvip_plain = decrypt(myvip_bytes, file_key) if myvip_bytes else None

        column_mappings = None
        if column_mappings_raw:
            import json
            try:
                column_mappings = json.loads(column_mappings_raw)
            except json.JSONDecodeError:
                raise AuditEngineError(
                    code="validation_failed",
                    message="column_mappings is not valid JSON",
                )

        inputs = AuditInputs(
            property_name=property_name,
            property_uuid=property_uuid,
            rent_roll=rent_roll_plain,
            rent_roll_filename=rent_roll_filename,
            vehicle_data=vehicle_plain,
            vehicle_data_filename=vehicle_data_filename,
            myvip_data=myvip_plain,
            myvip_data_filename=myvip_filename,
            column_mappings=column_mappings,
        )

        # 2) Run the audit (with hard timeout per contract §4.1)
        output = await asyncio.wait_for(
            asyncio.to_thread(run_audit, inputs),
            timeout=settings.max_processing_seconds,
        )

        # 3) Encrypt the workbook with the same per-audit key for the callback
        workbook_enc = encrypt(output.workbook_bytes, file_key)
        workbook_filename = f"{property_name.replace('/', '_')}_Parking_Audit.xlsx"
        metrics = output.metrics

    except asyncio.TimeoutError:
        error_code = "timeout"
        error_message = (
            f"Audit exceeded {settings.max_processing_seconds}s processing limit"
        )
        log.warning("[%s] timed out", audit_id)
    except AuditEngineError as exc:
        error_code = exc.code
        error_message = exc.message
        log.warning("[%s] engine error: %s — %s", audit_id, exc.code, exc.message)
    except Exception as exc:  # noqa: BLE001 — any unexpected failure → fail the audit
        error_code = "internal_error"
        error_message = "Audit service encountered an unexpected error."
        log.exception("[%s] unexpected error: %s", audit_id, exc)

    elapsed = time.monotonic() - started

    # 4) Update in-memory status (polling fallback)
    completed_at = datetime.now(timezone.utc)
    final_status = "completed" if error_code is None else "failed"
    async with _jobs_lock:
        _jobs[audit_id] = AuditStatusResponse(
            audit_id=audit_id,
            status=final_status,
            accepted_at=_jobs[audit_id].accepted_at if audit_id in _jobs else completed_at,
            completed_at=completed_at,
            error_code=error_code,
            error_message=error_message,
        )

    # 5) POST callback to dispatch (with retries per contract §8)
    await _post_callback(
        callback_url=callback_url,
        audit_id=audit_id,
        status=final_status,
        workbook_enc=workbook_enc,
        workbook_filename=workbook_filename,
        metrics=metrics,
        error_code=error_code,
        error_message=error_message,
        processed_seconds=elapsed,
    )

    async with _active_lock:
        _active_count -= 1


async def _post_callback(
    *,
    callback_url: str,
    audit_id: str,
    status: str,
    workbook_enc: Optional[bytes],
    workbook_filename: Optional[str],
    metrics: Optional[dict],
    error_code: Optional[str],
    error_message: Optional[str],
    processed_seconds: float,
) -> None:
    """POST the callback to dispatch with retry semantics from contract §8."""
    import json as _json

    token = issue_callback_token(audit_id)
    headers = {"Authorization": f"Bearer {token}"}

    files = []
    data = {
        "audit_id": audit_id,
        "status": status,
        "processed_seconds": f"{processed_seconds:.3f}",
    }
    if status == "completed" and workbook_enc is not None and workbook_filename:
        files.append(
            ("workbook", (workbook_filename, workbook_enc, "application/octet-stream"))
        )
        data["workbook_filename"] = workbook_filename
        if metrics is not None:
            data["metrics"] = _json.dumps(metrics)
    if status == "failed":
        data["error_code"] = error_code or "unknown"
        data["error_message"] = error_message or ""

    backoffs = [1, 4, 15, 60, 240]  # 5 attempts over ~5 minutes (contract §8)
    last_exc: Optional[Exception] = None
    async with httpx.AsyncClient(timeout=30.0) as client:
        for attempt, delay in enumerate(backoffs, start=1):
            try:
                resp = await client.post(callback_url, headers=headers, data=data, files=files)
                if 200 <= resp.status_code < 300:
                    log.info("[%s] callback delivered on attempt %d", audit_id, attempt)
                    return
                if 400 <= resp.status_code < 500:
                    # 4xx = no retry per contract §8
                    log.error(
                        "[%s] callback got %d — not retrying. body=%s",
                        audit_id, resp.status_code, resp.text[:500],
                    )
                    return
                log.warning(
                    "[%s] callback %d on attempt %d, will retry",
                    audit_id, resp.status_code, attempt,
                )
            except Exception as exc:  # noqa: BLE001
                last_exc = exc
                log.warning("[%s] callback attempt %d exception: %s", audit_id, attempt, exc)

            if attempt < len(backoffs):
                await asyncio.sleep(delay)

    log.error("[%s] callback abandoned after %d attempts (last error: %s)",
              audit_id, len(backoffs), last_exc)


# ------------------------------------------------------------------
# Hardening — never echo unexpected exceptions to clients
# ------------------------------------------------------------------
@app.exception_handler(Exception)
async def _unhandled_exception(request: Request, exc: Exception) -> JSONResponse:
    log.exception("unhandled error on %s: %s", request.url.path, exc)
    return JSONResponse(status_code=500, content={"error": "internal_error"})
