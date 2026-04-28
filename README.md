# Apollo Audit Service

Pure-Python microservice that runs apartment parking compliance audits for the Apollo Towing dispatch platform.

**Contract:** [`service-dashboard/docs/audit-service-contract.md`](../service-dashboard/docs/audit-service-contract.md) — APPROVED 2026-04-28. Any change here that deviates from the contract requires updating the doc first.

## Why this exists as a separate service

Apr 26, 2026 outage taught us: mixing Python (pdfplumber + openpyxl) into the dispatch Node.js Railway project breaks Railway's buildpack auto-detector and takes prod down. This service is fully decoupled — its own Railway project, its own buildpack, its own deploy lifecycle.

## Endpoints

| Method | Path | Auth | Purpose |
|--------|------|------|---------|
| `GET`  | `/v1/health` | none | Liveness probe |
| `POST` | `/v1/audits` | dispatch JWT | Hand off a job, returns 202 + processes async |
| `GET`  | `/v1/audits/{audit_id}` | none (lookup-only) | Polling fallback if dispatch never receives the callback |

The actual result is delivered via `POST <callback_url>` to dispatch — see contract §4.2.

## Run locally

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
# edit .env: set AUDIT_SERVICE_JWT_SECRET to a long random string
uvicorn app.main:app --reload --port 8000
```

Smoke test:
```bash
curl http://localhost:8000/v1/health
# {"status":"ok","version":"1.0.0","uptime_s":3}
```

## Deploy (Railway)

1. New Railway project, separate from dispatch.
2. Connect this repo.
3. Set env var `AUDIT_SERVICE_JWT_SECRET` — must match the value on the dispatch Railway project exactly.
4. Railway auto-detects Python via Nixpacks (no `nixpacks.toml`, no Dockerfile — that was the Apr 26 trap).
5. `Procfile` provides the start command. `$PORT` is injected by Railway.
6. Add custom domain `audit.apollotowingdfw.com`.

## Project layout

```
audit-service/
  app/
    main.py          FastAPI app + the 3 endpoints + background processing
    auth.py          HS256 JWT verify (incoming) + sign (callback)
    crypto.py        AES-256-GCM helpers (per-audit key from JWT claim)
    audit_engine.py  STUB — port target for the permit-audit skill
    models.py        Pydantic request/response shapes
    settings.py      pydantic-settings env config
  requirements.txt
  Procfile           web: uvicorn app.main:app --host 0.0.0.0 --port $PORT
  .env.example
  .gitignore
```

## What's NOT implemented yet

- `app/audit_engine.run_audit()` is a stub that raises `not_implemented`. The real port from the `permit-audit` skill (rent-roll parsing via pdfplumber, vehicle-data normalization, name matching, openpyxl workbook generation) is the next deliverable.
- Tests. Pytest suite + a fixture audit (one anonymized rent roll + vehicle file + expected workbook) lands with the engine port.

## Operational notes

- **Stateless** — no DB, no persistent volume. In-memory job table powers the `/v1/audits/{id}` polling fallback only; lost on restart by design (dispatch's 15-min callback timeout catches orphans).
- **Concurrency cap** — `MAX_CONCURRENT_AUDITS=3` by default. Returns 503 above that.
- **Hard timeout per audit** — `MAX_PROCESSING_SECONDS=120` (contract §4.1). Beyond that, the job aborts and posts `failed` to the callback.
- **Logging** — `audit_id` + processing time + error class only. Never logs file contents or `file_key`.
