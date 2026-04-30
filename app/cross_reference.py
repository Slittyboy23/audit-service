"""Cross-reference rent roll vs vehicle data → compliance audit data model.

Ported from the permit-audit skill. The output of this module feeds directly
into the workbook builder (5 sheets defined in SKILL.md).
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from datetime import date
from typing import Dict, List, Optional

from .normalize import extract_last_name, is_skip_name
from .parsers import RentRollEntry, Vehicle


# Audit status values per SKILL.md Sheet 2
STATUS_REGISTERED = "Registered"
STATUS_TOW_RISK = "TOW RISK"
STATUS_VACANT_HAS_PERMITS = "Vacant - Has Permits"
STATUS_VACANT_OK = "Vacant - OK"
STATUS_PARKING_ONLY = "Parking Only"


@dataclass
class CrossRefRow:
    apt_canonical: str
    apt_display: str
    name: str
    occupancy: str            # 'Occupied' | 'Vacant' | 'Model' | 'NTV'
    in_parking_system: bool
    vehicles: int
    audit_status: str         # one of the STATUS_* constants


@dataclass
class VacantRow:
    apt_display: str
    status_label: str         # 'Vacant' | 'Model' | 'Future Resident' | 'NTV'
    name: str
    in_parking_system: bool
    vehicles: int


@dataclass
class DiscrepancyRow:
    apt_display: str
    rent_roll_name: str
    permit_name: str


@dataclass
class AuditData:
    property_name: str
    rent_roll_date: Optional[str]
    cross_ref: List[CrossRefRow] = field(default_factory=list)
    vehicles: List[Vehicle] = field(default_factory=list)
    vacants: List[VacantRow] = field(default_factory=list)
    discrepancies: List[DiscrepancyRow] = field(default_factory=list)
    rent_roll_units: int = 0
    occupied: int = 0
    vacant: int = 0
    model: int = 0
    ntv: int = 0
    compliance_rate: float = 0.0
    tow_risk: int = 0
    total_vehicles: int = 0
    # ─── Cover-sheet enrichment from wizard Step 1 (optional) ───
    property_address: str = ""
    audit_date: Optional[str] = None         # ISO yyyy-mm-dd, falls back to today
    parking_program: Optional[str] = None    # 'myvip' | 'permit'
    num_units_form: int = 0                  # what the user typed in Step 1
    parking_lot: dict = field(default_factory=dict)  # reserved/open/guest/handicap/specialty/total
    specialty_tags: List[str] = field(default_factory=list)
    # Distribution of registered vehicles across occupied units (and overall).
    # Keyed by bucket label: '0', '1', '2', '3', '4', '5', '6+'.
    vehicle_distribution: dict = field(default_factory=dict)
    # MyVIP profile coverage — only meaningful when parking_program == 'myvip'.
    # Pulled from form_data.myvip_summary.myvip_profile_units. The cover renders
    # `Units Without MyVIP Profile = occupied - myvip_profile_units` as a TOW
    # RISK callout.
    myvip_profile_units: int = 0


_OCCUPANCY_LABEL = {
    "occupied": "Occupied",
    "vacant": "Vacant",
    "model": "Model",
    "ntv": "NTV",
}


def _display_apt(canonical: str, raw: str) -> str:
    """Preserve the property's display style. If raw had leading zeros, keep them."""
    raw_clean = (raw or "").lstrip("'").strip()
    digits = "".join(c for c in raw_clean if c.isdigit())
    if digits and digits != canonical and digits.lstrip("0") == canonical:
        return digits
    return canonical


def cross_reference(
    *,
    property_name: str,
    rent_roll: List[RentRollEntry],
    vehicles: List[Vehicle],
    rent_roll_date: Optional[str] = None,
    form_data: Optional[dict] = None,
) -> AuditData:
    """Produce the AuditData model that drives the 5-sheet workbook."""

    # Index vehicles by canonical apartment number
    by_apt: Dict[str, List[Vehicle]] = defaultdict(list)
    for v in vehicles:
        by_apt[v.apt].append(v)

    # Index rent roll: take FIRST occurrence as the current resident,
    # additional occurrences flagged as future residents.
    primary: Dict[str, RentRollEntry] = {}
    future: List[RentRollEntry] = []
    for e in rent_roll:
        if e.is_future:
            future.append(e)
        else:
            primary[e.apt] = e

    data = AuditData(property_name=property_name, rent_roll_date=rent_roll_date)

    # ----- Sheet 2: Cross-reference rows ----------------------------------
    rent_roll_apts = sorted(primary.keys(), key=_apt_sort_key)
    for apt in rent_roll_apts:
        entry = primary[apt]
        veh_list = by_apt.get(apt, [])
        in_system = bool(veh_list)
        veh_count = len(veh_list)
        status = _classify_audit_status(entry, in_system, veh_count)
        data.cross_ref.append(CrossRefRow(
            apt_canonical=apt,
            apt_display=_display_apt(apt, entry.raw_apt),
            name=entry.name,
            occupancy=_OCCUPANCY_LABEL.get(entry.status, entry.status.title()),
            in_parking_system=in_system,
            vehicles=veh_count,
            audit_status=status,
        ))

    # Parking-only units (no rent-roll match)
    parking_only_apts = sorted(
        [a for a in by_apt.keys() if a not in primary],
        key=_apt_sort_key,
    )
    for apt in parking_only_apts:
        veh_list = by_apt[apt]
        sample = veh_list[0]
        data.cross_ref.append(CrossRefRow(
            apt_canonical=apt,
            apt_display=_display_apt(apt, sample.raw_apt),
            name="(not on rent roll)",
            occupancy="—",
            in_parking_system=True,
            vehicles=len(veh_list),
            audit_status=STATUS_PARKING_ONLY,
        ))

    # ----- Sheet 3: Vehicle details ---------------------------------------
    # Output every vehicle, plus a NONE row for occupied units with no vehicles.
    sorted_vehicles: List[Vehicle] = []
    apts_with_no_vehicles_to_emit: List[str] = []
    for apt in rent_roll_apts:
        if by_apt.get(apt):
            sorted_vehicles.extend(by_apt[apt])
        else:
            entry = primary[apt]
            if entry.status == "occupied":
                apts_with_no_vehicles_to_emit.append(apt)
    for apt in parking_only_apts:
        sorted_vehicles.extend(by_apt[apt])
    data.vehicles = sorted_vehicles
    # The "NONE" rows are signaled by storing empty Vehicle records the workbook
    # builder can recognize via apt_canonical present + plate empty + name empty.
    # Simpler: pass them through a side channel.
    data._no_vehicle_apts = apts_with_no_vehicles_to_emit  # type: ignore[attr-defined]

    # ----- Sheet 4: Vacant & Future ---------------------------------------
    vacant_rows: List[VacantRow] = []
    for apt in rent_roll_apts:
        e = primary[apt]
        if e.status in {"vacant", "model", "ntv"}:
            vacant_rows.append(VacantRow(
                apt_display=_display_apt(apt, e.raw_apt),
                status_label={"vacant": "Vacant", "model": "Model", "ntv": "NTV"}[e.status],
                name=e.name if not is_skip_name(e.name) else "",
                in_parking_system=apt in by_apt,
                vehicles=len(by_apt.get(apt, [])),
            ))
    for fut in future:
        vacant_rows.append(VacantRow(
            apt_display=_display_apt(fut.apt, fut.raw_apt),
            status_label="Future Resident",
            name=fut.name,
            in_parking_system=fut.apt in by_apt,
            vehicles=len(by_apt.get(fut.apt, [])),
        ))
    data.vacants = vacant_rows

    # ----- Sheet 5: Name discrepancies ------------------------------------
    for apt, e in primary.items():
        if e.status != "occupied" or is_skip_name(e.name):
            continue
        rent_last = extract_last_name(e.name)
        if not rent_last:
            continue
        veh_list = by_apt.get(apt, [])
        if not veh_list:
            continue
        # Discrepancy if NONE of the registered vehicles' name fields match.
        permit_lasts = {extract_last_name(v.name) for v in veh_list if v.name and not is_skip_name(v.name)}
        permit_lasts.discard("")
        if permit_lasts and rent_last not in permit_lasts:
            data.discrepancies.append(DiscrepancyRow(
                apt_display=_display_apt(apt, e.raw_apt),
                rent_roll_name=e.name,
                permit_name=", ".join(sorted({v.name for v in veh_list if v.name})),
            ))

    # ----- Stats block (Sheet 1) ------------------------------------------
    data.rent_roll_units = len(primary)
    for e in primary.values():
        if e.status == "occupied":
            data.occupied += 1
        elif e.status == "vacant":
            data.vacant += 1
        elif e.status == "model":
            data.model += 1
        elif e.status == "ntv":
            data.ntv += 1
    data.total_vehicles = len(vehicles)
    occupied_with_vehicles = sum(
        1 for apt, e in primary.items()
        if e.status == "occupied" and by_apt.get(apt)
    )
    data.compliance_rate = (occupied_with_vehicles / data.occupied) if data.occupied else 0.0
    data.tow_risk = sum(
        1 for apt, e in primary.items()
        if e.status == "occupied" and not by_apt.get(apt)
    )

    # ----- Vehicle distribution buckets (cover-sheet histogram) -----------
    # Bucket every OCCUPIED unit by how many vehicles it has registered.
    # 0..5 are exact; 6+ is the overflow bucket. Always emit all keys so the
    # cover sheet can render a complete histogram.
    buckets = {"0": 0, "1": 0, "2": 0, "3": 0, "4": 0, "5": 0, "6+": 0}
    for apt, e in primary.items():
        if e.status != "occupied":
            continue
        n = len(by_apt.get(apt, []))
        key = str(n) if n <= 5 else "6+"
        buckets[key] += 1
    data.vehicle_distribution = buckets

    # ----- Optional form-data enrichment (cover sheet) --------------------
    if form_data:
        addr = form_data.get("property_address") or ""
        data.property_address = addr.strip()

        ad = form_data.get("audit_date")
        data.audit_date = ad if isinstance(ad, str) and ad else None

        pp = form_data.get("parking_program")
        if pp in ("myvip", "permit"):
            data.parking_program = pp

        try:
            data.num_units_form = int(form_data.get("units") or 0)
        except (TypeError, ValueError):
            data.num_units_form = 0

        # Parking lot space breakdown — every key independent
        def _int(k):
            try:
                return int(form_data.get(k) or 0)
            except (TypeError, ValueError):
                return 0
        reserved = _int("reserved")
        open_spc = _int("open_spaces")
        guest = _int("guest")
        handicap = _int("handicap")
        specialty = _int("specialty")
        data.parking_lot = {
            "reserved": reserved,
            "open": open_spc,
            "guest": guest,
            "handicap": handicap,
            "specialty": specialty,
            "total": reserved + open_spc + guest + handicap + specialty,
        }

        tags = form_data.get("specialty_tags")
        if isinstance(tags, list):
            data.specialty_tags = [str(t) for t in tags if t]

        # MyVIP profile coverage (Step 4 wizard panel feeds this)
        mv = form_data.get("myvip_summary") or {}
        try:
            data.myvip_profile_units = int(mv.get("myvip_profile_units") or 0)
        except (TypeError, ValueError):
            data.myvip_profile_units = 0

    return data


def _apt_sort_key(apt: str):
    """Sort apartments numerically, falling back to lexical for non-digit codes."""
    try:
        return (0, int(apt))
    except ValueError:
        return (1, apt)


def _classify_audit_status(entry: RentRollEntry, in_system: bool, veh_count: int) -> str:
    if entry.status == "occupied":
        return STATUS_REGISTERED if veh_count > 0 else STATUS_TOW_RISK
    if entry.status in {"vacant", "ntv"}:
        return STATUS_VACANT_HAS_PERMITS if veh_count > 0 else STATUS_VACANT_OK
    if entry.status == "model":
        return STATUS_VACANT_HAS_PERMITS if veh_count > 0 else STATUS_VACANT_OK
    return STATUS_VACANT_OK
