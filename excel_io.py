"""Phase 5c - Excel I/O for MS Project hakedis workflow.

Pure-Python openpyxl helpers. MSP/COM/file independent - takes plain
dicts, writes openpyxl Workbook objects. MCS brand styling per CLAUDE.md
RULE 14.
"""
from openpyxl.styles import Font, PatternFill, Alignment


# ---------- MCS brand constants (CLAUDE.md RULE 14) ----------

BRAND_LACIVERT = "FF0B1F4D"   # ARGB - main brand
BRAND_NAVY = "FF3D4663"
BRAND_TURQUOISE = "FF39B4CC"
BRAND_LABEL_GRAY = "FF6B7394"
ZEBRA = "FFF0F3F8"

# RAG status colors (industry standard, no competitor names per RULE 13)
RAG_GREEN = "FF8FBC8F"
RAG_AMBER = "FFFFCC66"
RAG_RED = "FFE57373"

WHITE = "FFFFFFFF"


def apply_header_style(ws, row=1):
    """Lacivert background + white bold Calibri 11pt header (CLAUDE.md RULE 14)."""
    fill = PatternFill(start_color=BRAND_LACIVERT, end_color=BRAND_LACIVERT,
                       fill_type="solid")
    font = Font(name="Calibri", bold=True, size=11, color=WHITE)
    align = Alignment(horizontal="center", vertical="center")
    for cell in ws[row]:
        cell.fill = fill
        cell.font = font
        cell.alignment = align


def apply_rag_fill(cell, status):
    """Apply RAG color fill based on status string (pass/fail/amber/green/red)."""
    s = (status or "").lower()
    if s in ("pass", "green", "ok"):
        color = RAG_GREEN
    elif s in ("fail", "red"):
        color = RAG_RED
    elif s == "amber":
        color = RAG_AMBER
    else:
        return
    cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")


def apply_zebra_fill(cell, row_idx):
    """Light gray zebra stripe on even rows (1-indexed). Odd rows untouched."""
    if row_idx % 2 == 0:
        cell.fill = PatternFill(start_color=ZEBRA, end_color=ZEBRA, fill_type="solid")
