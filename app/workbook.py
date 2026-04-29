"""5-sheet audit workbook builder.

Implements the exact formatting spec from the permit-audit skill (SKILL.md
"Formatting Standards" section). All colors, fonts, fills, borders, and
column widths are taken verbatim from the spec.
"""
from __future__ import annotations

import io
from datetime import date
from typing import Iterable

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .cross_reference import (
    AuditData,
    STATUS_PARKING_ONLY,
    STATUS_REGISTERED,
    STATUS_TOW_RISK,
    STATUS_VACANT_HAS_PERMITS,
    STATUS_VACANT_OK,
)


# ------------------------------------------------------------------
# Color palette — straight from SKILL.md
# ------------------------------------------------------------------
DARK_BLUE = "2F5496"
MEDIUM_BLUE = "4472C4"
GREEN_FILL = "C6EFCE"
GREEN_TEXT = "006100"
RED_FILL = "FFC7CE"
RED_TEXT = "9C0006"
YELLOW_FILL = "FFEB9C"
YELLOW_TEXT = "9C6500"
ORANGE_HIGHLIGHT = "FFF2CC"
ALT_ROW = "F2F2F2"
NAVY = "1F4E79"
BORDER_GRAY = "D9D9D9"
RED_BRIGHT = "FF0000"
RED_DEEP = "C00000"
SUBTITLE_GRAY = "404040"
GENERATED_GRAY = "808080"

# ------------------------------------------------------------------
# Reusable style fragments
# ------------------------------------------------------------------
ARIAL = "Arial"
THIN = Side(style="thin", color=BORDER_GRAY)
DATA_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WHITE = "FFFFFF"


def _header_font() -> Font:
    return Font(name=ARIAL, size=11, bold=True, color=WHITE)


def _data_font() -> Font:
    return Font(name=ARIAL, size=10)


def _fill(color: str) -> PatternFill:
    return PatternFill(start_color=color, end_color=color, fill_type="solid")


def _set_columns(ws: Worksheet, widths: Iterable[int]) -> None:
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _apply_header_row(ws: Worksheet, headers: Iterable[str], fill_color: str) -> None:
    for col, value in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=value)
        cell.fill = _fill(fill_color)
        cell.font = _header_font()
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = DATA_BORDER


def _apply_alt_shading(ws: Worksheet, start_row: int, end_row: int, ncols: int) -> None:
    for r in range(start_row, end_row + 1):
        if (r - start_row) % 2 == 1:
            for c in range(1, ncols + 1):
                ws.cell(row=r, column=c).fill = _fill(ALT_ROW)


def _set_data_borders(ws: Worksheet, start_row: int, end_row: int, ncols: int) -> None:
    for r in range(start_row, end_row + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = DATA_BORDER
            if not cell.font or cell.font.name is None:
                cell.font = _data_font()


# ------------------------------------------------------------------
# Public entry point
# ------------------------------------------------------------------
def build_workbook(audit: AuditData) -> bytes:
    wb = Workbook()
    # Drop the default sheet — we make our own
    default = wb.active
    wb.remove(default)

    _build_summary(wb, audit)
    _build_cross_reference(wb, audit)
    _build_vehicle_details(wb, audit)
    _build_vacant_future(wb, audit)
    _build_discrepancies(wb, audit)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ------------------------------------------------------------------
# Sheet 1 — Audit Summary
# ------------------------------------------------------------------
def _build_summary(wb: Workbook, audit: AuditData) -> None:
    ws = wb.create_sheet("Audit Summary")
    ws.sheet_properties.tabColor = DARK_BLUE

    ws.cell(row=1, column=1, value=audit.property_name).font = Font(
        name=ARIAL, size=14, bold=True, color=DARK_BLUE,
    )
    rr_date = audit.rent_roll_date or date.today().strftime("%m/%d/%Y")
    ws.cell(row=2, column=1, value=f"Rent Roll ({rr_date}) vs. Parking Software Permits/Vehicles").font = Font(
        name=ARIAL, size=11, color=SUBTITLE_GRAY,
    )
    ws.cell(row=3, column=1, value=f"Generated {date.today().strftime('%m/%d/%Y')}").font = Font(
        name=ARIAL, size=10, italic=True, color=GENERATED_GRAY,
    )

    rows = [
        ("Total Units on Rent Roll", audit.rent_roll_units),
        ("Occupied Units", audit.occupied),
        ("Vacant Units", audit.vacant),
        ("Model Units", audit.model),
        ("Units with Notice to Vacate", audit.ntv),
        ("Compliance Rate", f"{audit.compliance_rate * 100:.1f}%"),
        ("TOW RISK (occupied, no vehicles)", audit.tow_risk),
        ("Total Vehicles Registered", audit.total_vehicles),
        ("Name Discrepancies", len(audit.discrepancies)),
    ]
    start = 5
    for i, (label, value) in enumerate(rows):
        r = start + i
        lbl = ws.cell(row=r, column=1, value=label)
        val = ws.cell(row=r, column=2, value=value)
        lbl.font = Font(name=ARIAL, size=11, bold=True)
        val.font = Font(name=ARIAL, size=11)
        if label.startswith("TOW RISK") and audit.tow_risk > 0:
            val.font = Font(name=ARIAL, size=11, bold=True, color=RED_BRIGHT)
        lbl.border = DATA_BORDER
        val.border = DATA_BORDER

    _set_columns(ws, [38, 18])


# ------------------------------------------------------------------
# Sheet 2 — Full Cross-Reference
# ------------------------------------------------------------------
def _build_cross_reference(wb: Workbook, audit: AuditData) -> None:
    ws = wb.create_sheet("Cross-Reference")
    ws.sheet_properties.tabColor = MEDIUM_BLUE

    headers = [
        "Unit #", "Resident Name", "Occupancy Status",
        "In Parking System?", "# Vehicles Registered", "Audit Status",
    ]
    _apply_header_row(ws, headers, DARK_BLUE)

    status_styles = {
        STATUS_REGISTERED: (GREEN_FILL, Font(name=ARIAL, size=10, color=GREEN_TEXT)),
        STATUS_TOW_RISK: (RED_FILL, Font(name=ARIAL, size=10, bold=True, color=RED_TEXT)),
        STATUS_VACANT_HAS_PERMITS: (YELLOW_FILL, Font(name=ARIAL, size=10, color=YELLOW_TEXT)),
        STATUS_VACANT_OK: (YELLOW_FILL, Font(name=ARIAL, size=10, color=YELLOW_TEXT)),
        STATUS_PARKING_ONLY: (None, Font(name=ARIAL, size=10)),
    }

    for i, row in enumerate(audit.cross_ref):
        r = i + 2
        ws.cell(row=r, column=1, value=row.apt_display).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=row.name)
        ws.cell(row=r, column=3, value=row.occupancy)
        ws.cell(row=r, column=4, value="Yes" if row.in_parking_system else "No")
        ws.cell(row=r, column=5, value=row.vehicles).alignment = Alignment(horizontal="center")
        status_cell = ws.cell(row=r, column=6, value=row.audit_status)
        fill_color, font = status_styles.get(row.audit_status, (None, _data_font()))
        if fill_color:
            status_cell.fill = _fill(fill_color)
        status_cell.font = font

    end = 1 + len(audit.cross_ref)
    _set_data_borders(ws, 2, end, len(headers))
    _set_columns(ws, [10, 30, 18, 18, 22, 22])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(end, 2)}"


# ------------------------------------------------------------------
# Sheet 3 — Vehicle Details by Unit
# ------------------------------------------------------------------
def _build_vehicle_details(wb: Workbook, audit: AuditData) -> None:
    ws = wb.create_sheet("Vehicle Details")
    ws.sheet_properties.tabColor = "70AD47"

    headers = [
        "Unit #", "Resident (Rent Roll)", "Registered Name",
        "Make", "Model", "License Plate", "Date Registered",
    ]
    _apply_header_row(ws, headers, "70AD47")

    # Build a quick lookup of rent-roll names by canonical apt
    rr_name = {row.apt_canonical: row.name for row in audit.cross_ref}

    r = 2
    for v in audit.vehicles:
        ws.cell(row=r, column=1, value=v.raw_apt or v.apt).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=rr_name.get(v.apt, ""))
        ws.cell(row=r, column=3, value=v.name)
        ws.cell(row=r, column=4, value=v.make)
        ws.cell(row=r, column=5, value=v.model)
        ws.cell(row=r, column=6, value=v.plate)
        ws.cell(row=r, column=7, value=v.date_registered)
        r += 1

    # NONE rows for occupied units with zero vehicles
    none_apts = getattr(audit, "_no_vehicle_apts", []) or []
    cross_by_apt = {row.apt_canonical: row for row in audit.cross_ref}
    none_font = Font(name=ARIAL, size=10, bold=True, color=RED_BRIGHT)
    for apt in none_apts:
        cr = cross_by_apt.get(apt)
        if not cr:
            continue
        ws.cell(row=r, column=1, value=cr.apt_display).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=cr.name)
        for col in (3, 4, 5, 6, 7):
            cell = ws.cell(row=r, column=col, value="NONE")
            cell.font = none_font
        r += 1

    end = r - 1
    _set_data_borders(ws, 2, end, len(headers))
    _set_columns(ws, [10, 28, 28, 12, 14, 16, 18])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(end, 2)}"


# ------------------------------------------------------------------
# Sheet 4 — Vacant & Future Residents
# ------------------------------------------------------------------
def _build_vacant_future(wb: Workbook, audit: AuditData) -> None:
    ws = wb.create_sheet("Vacant & Future")
    ws.sheet_properties.tabColor = "FFC000"

    headers = ["Unit #", "Status", "Resident Name", "In Parking System?", "# Vehicles"]
    _apply_header_row(ws, headers, "FFC000")

    for i, v in enumerate(audit.vacants):
        r = i + 2
        ws.cell(row=r, column=1, value=v.apt_display).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=v.status_label)
        ws.cell(row=r, column=3, value=v.name)
        ws.cell(row=r, column=4, value="Yes" if v.in_parking_system else "No")
        ws.cell(row=r, column=5, value=v.vehicles).alignment = Alignment(horizontal="center")

    end = 1 + len(audit.vacants)
    _set_data_borders(ws, 2, end, len(headers))
    _set_columns(ws, [10, 18, 28, 20, 14])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(end, 2)}"


# ------------------------------------------------------------------
# Sheet 5 — Name Discrepancies
# ------------------------------------------------------------------
def _build_discrepancies(wb: Workbook, audit: AuditData) -> None:
    ws = wb.create_sheet("Name Discrepancies")
    ws.sheet_properties.tabColor = "FF6600"

    ws.cell(row=1, column=1, value=f"{audit.property_name} — NAME DISCREPANCIES").font = Font(
        name=ARIAL, size=14, bold=True,
    )
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)

    subtitle = (
        f"Units where Rent Roll name does not match Parking Permit name | "
        f"{len(audit.discrepancies)} discrepancies found"
    )
    ws.cell(row=2, column=1, value=subtitle).font = Font(name=ARIAL, size=10, color="666666")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=3)

    headers = ["Unit", "Rent Roll Name", "Permit/Profile Name"]
    for col, value in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=col, value=value)
        cell.fill = _fill(NAVY)
        cell.font = _header_font()
        cell.alignment = Alignment(horizontal="center")
        cell.border = DATA_BORDER

    for i, d in enumerate(audit.discrepancies):
        r = 5 + i
        ws.cell(row=r, column=1, value=d.apt_display).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=d.rent_roll_name)
        permit_cell = ws.cell(row=r, column=3, value=d.permit_name)
        permit_cell.fill = _fill(ORANGE_HIGHLIGHT)
        for c in range(1, 4):
            cell = ws.cell(row=r, column=c)
            cell.border = DATA_BORDER
            if not cell.font.name:
                cell.font = _data_font()

    # Total at bottom
    total_row = 5 + len(audit.discrepancies) + 1
    total_cell = ws.cell(
        row=total_row, column=1,
        value=f"Total discrepancies: {len(audit.discrepancies)}",
    )
    total_cell.font = Font(name=ARIAL, size=11, bold=True, color=RED_DEEP)
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=3)

    _set_columns(ws, [10, 32, 32])
    ws.freeze_panes = "A5"
