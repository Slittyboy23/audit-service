"""Audit engine — port target for the `permit-audit` skill.

STATUS: STUB. The real port from the permit-audit skill (pdfplumber rent-roll
parsing, vehicle-data normalization, name-matching, openpyxl workbook gen)
lands in a follow-up commit.

Contract: this module is called from app.main with the decrypted file bytes
plus the property metadata, and must return:
  - workbook_bytes: the unencrypted .xlsx as bytes
  - metrics: dict matching app.models.AuditMetrics

If parsing fails, raise AuditEngineError with a machine-readable code +
human-readable message. Per contract §11, v1 is all-or-nothing — partial
results are explicitly out of scope.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


class AuditEngineError(Exception):
    """Raised on any audit failure. `code` flows to dispatch as `error_code`."""

    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.code = code
        self.message = message


@dataclass
class AuditInputs:
    property_name: str
    property_uuid: str
    rent_roll: bytes
    rent_roll_filename: str
    vehicle_data: bytes
    vehicle_data_filename: str
    myvip_data: Optional[bytes] = None
    myvip_data_filename: Optional[str] = None
    column_mappings: Optional[dict] = None


@dataclass
class AuditOutput:
    workbook_bytes: bytes
    metrics: dict


def run_audit(inputs: AuditInputs) -> AuditOutput:
    """Run the parking compliance audit end-to-end.

    TODO: port from permit-audit skill. For now, this raises so the service
    never silently returns empty results in prod.
    """
    raise AuditEngineError(
        code="not_implemented",
        message="Audit engine port from permit-audit skill is pending.",
    )
