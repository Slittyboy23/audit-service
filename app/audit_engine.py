"""Audit engine — orchestrates parsing, cross-reference, and workbook generation.

Faithful port of the `permit-audit` skill (SKILL.md). Errors raised as
AuditEngineError flow through to dispatch as `error_code` + `error_message`
(contract §4.2).
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
    # Wizard Step 1 form payload — property address, parking program,
    # parking-lot space breakdown, specialty tags. Used for the marketing
    # cover sheet (Sheet 1). Optional — when absent the cover falls back
    # to whatever it can compute from the rent roll + vehicles alone.
    form_data: Optional[dict] = None


@dataclass
class AuditOutput:
    workbook_bytes: bytes
    metrics: dict


def run_audit(inputs: AuditInputs) -> AuditOutput:
    """Run the parking compliance audit end-to-end.

    1. Parse rent roll (PDF or XLSX).
    2. Parse vehicle/permit data (CSV or XLSX).
    3. Cross-reference (compliance, tow risk, name discrepancies).
    4. Generate the 5-sheet workbook.
    5. Return metrics matching contract §7.
    """
    # Imports are deferred to keep this module light at startup
    from .cross_reference import cross_reference
    from .parsers import parse_rent_roll, parse_vehicle_data
    from .workbook import build_workbook

    rent_roll = parse_rent_roll(inputs.rent_roll, inputs.rent_roll_filename)
    vehicles = parse_vehicle_data(inputs.vehicle_data, inputs.vehicle_data_filename)

    audit = cross_reference(
        property_name=inputs.property_name,
        rent_roll=rent_roll,
        vehicles=vehicles,
        rent_roll_date=None,  # could be extracted from PDF first page later
        form_data=inputs.form_data,
    )

    workbook_bytes = build_workbook(audit)

    # Detect vehicle data source from filename — feeds Sheet 1 subtitle and metrics
    src = inputs.vehicle_data_filename or ""
    if "permitclick" in src.lower():
        vehicle_source = "PermitClick"
    elif "myvip" in src.lower() or "profile" in src.lower():
        vehicle_source = "MyVIP"
    else:
        vehicle_source = "Parking System"

    metrics = {
        "total_units": audit.rent_roll_units,
        "occupied_units": audit.occupied,
        "vacant_units": audit.vacant,
        "model_units": audit.model,
        "ntv_units": audit.ntv,
        "compliance_rate": round(audit.compliance_rate, 4),
        "tow_risk_units": audit.tow_risk,
        "total_vehicles": audit.total_vehicles,
        "name_discrepancies": len(audit.discrepancies),
        "rent_roll_date": audit.rent_roll_date,
        "vehicle_data_source": vehicle_source,
    }
    return AuditOutput(workbook_bytes=workbook_bytes, metrics=metrics)
