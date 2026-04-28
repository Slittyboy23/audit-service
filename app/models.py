"""Pydantic models for API request/response shapes (contract §4, §7)."""
from __future__ import annotations

from datetime import datetime
from typing import Optional

from pydantic import BaseModel, Field


class AcceptedResponse(BaseModel):
    audit_id: str
    accepted_at: datetime
    estimated_seconds: int = 45


class HealthResponse(BaseModel):
    status: str = "ok"
    version: str
    uptime_s: int


class AuditMetrics(BaseModel):
    """Workbook-derived numbers stored on dispatch's `audits` row (contract §7)."""

    total_units: int
    occupied_units: int
    vacant_units: int
    model_units: int = 0
    ntv_units: int = 0
    compliance_rate: float = Field(ge=0, le=1)
    tow_risk_units: int
    total_vehicles: int
    name_discrepancies: int
    rent_roll_date: Optional[str] = None  # ISO date string
    vehicle_data_source: Optional[str] = None


class AuditStatusResponse(BaseModel):
    """Returned by GET /v1/audits/{id} (polling fallback)."""

    audit_id: str
    status: str  # queued | processing | completed | failed
    accepted_at: Optional[datetime] = None
    completed_at: Optional[datetime] = None
    error_code: Optional[str] = None
    error_message: Optional[str] = None
