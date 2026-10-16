"""Per-program settings. The card engine is shared; only what genuinely differs
between D240 and D53 lives here. Adding a third program = adding one entry."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_DIR = APP_DIR / "templates"
FONT_DIR = APP_DIR / "fonts"


@dataclass(frozen=True)
class ProgramConfig:
    code: str                      # "D240" / "D53"
    title: str                     # tab title
    template_file: str             # bundled Target template (Instructions/Template/Dropdowns)
    # How each card's "Program Name" (card row 3, and the IMU-goal lookup key) is found:
    #   "subclass" -> "Subclass <Subclass ID>" from the PCN row   (D240)
    #   "filename" -> the part of the PCN file name after " - "   (D53)
    program_source: str
    default_yellow_band: float     # used only if the Dropdowns sheet has no band formula
    notes: tuple[str, ...] = field(default_factory=tuple)

    @property
    def template_path(self) -> Path:
        return TEMPLATE_DIR / self.template_file


PROGRAMS: dict[str, ProgramConfig] = {
    "D240": ProgramConfig(
        code="D240",
        title="D240",
        template_file="D240_Sample_Tag_Template.xlsx",
        program_source="subclass",
        default_yellow_band=0.04,
        notes=("Sample Tag 上的 Subclass 依 PCN 的 Subclass ID 帶出（例：Subclass 27）。",),
    ),
    "D53": ProgramConfig(
        code="D53",
        title="D53",
        template_file="D53_Sample_Tag_Template.xlsx",
        program_source="subclass",
        default_yellow_band=0.04,
        notes=("Sample Tag 上的 Subclass 依 PCN 的 Subclass ID 帶出（例：Subclass 23）。",),
    ),
}

# Card geometry shared by every Target sample-tag template seen so far
# (D240 2026, D240 27C5, D53 27C5): 25-row blocks starting at row 2,
# left card in B:F, right card in H:L.
BLOCK_START = 2
BLOCK_ROWS = 25          # rows 2..26 (row 26 is the spacer)
CARD_ROWS = 24           # rows 2..25 are the visible card
LEFT_COL = 2             # B
RIGHT_COL = 8            # H
CARD_COLS = 5            # B..F
LAST_COL = 13            # M (print area A:M)
BLOCKS_PER_PAGE = 4      # print: same as Target template (rows 1-101 = 4 blocks = 8 cards per page)
