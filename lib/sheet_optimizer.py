"""
2D sheet goods optimizer using the Guillotine cut algorithm.

Pure Python, no Fusion 360 API dependencies. All dimensions are in inches.
"""

from dataclasses import dataclass, field


DEFAULT_SHEET_SIZE = '48x96'   # 4×8 plywood / OSB


@dataclass
class PlacedPart:
    """A part that has been placed on a sheet."""

    x: float
    y: float
    width: float
    height: float
    label: str
    rotated: bool = False


@dataclass
class FreeRect:
    """An available rectangle of space on a sheet."""

    x: float
    y: float
    width: float
    height: float


@dataclass
class Sheet:
    """A single sheet with its placed parts and remaining free rectangles."""

    sheet_width: float
    sheet_height: float
    placed: list = field(default_factory=list)
    _free_rects: list = field(init=False, default_factory=list)

    def __post_init__(self):
        self._free_rects = [FreeRect(0.0, 0.0, self.sheet_width, self.sheet_height)]

    @property
    def used_area(self) -> float:
        return sum(p.width * p.height for p in self.placed)

    @property
    def total_area(self) -> float:
        return self.sheet_width * self.sheet_height

    @property
    def waste_area(self) -> float:
        return self.total_area - self.used_area

    @property
    def waste_pct(self) -> float:
        return (self.waste_area / self.total_area * 100) if self.total_area > 0 else 0.0


@dataclass
class SheetPlan:
    """Result of a sheet optimization run."""

    sheets: list
    sheet_width: float
    sheet_height: float
    oversized: list = field(default_factory=list)

    @property
    def total_waste_area(self) -> float:
        return sum(s.waste_area for s in self.sheets)

    @property
    def total_area(self) -> float:
        return len(self.sheets) * self.sheet_width * self.sheet_height

    @property
    def waste_pct(self) -> float:
        return (self.total_waste_area / self.total_area * 100) if self.total_area > 0 else 0.0


def parse_sheet_size(value: str):
    """
    Parse a sheet size string into a (width, height) tuple of floats.

    Accepts '48x96', '48X96', '48, 96', etc.
    Returns None if the string is blank or cannot be parsed.
    """
    if not value or not value.strip():
        return None
    for sep in ('x', 'X', ','):
        tokens = value.split(sep, 1)
        if len(tokens) == 2:
            try:
                w = float(tokens[0].strip())
                h = float(tokens[1].strip())
                if w > 0 and h > 0:
                    return (w, h)
            except ValueError:
                continue
    return None


def _best_fit(free_rects, pw, ph, allow_rotation):
    """
    Find the free rect that best fits the part using Best Short Side Fit.

    Returns (rect, rotated) or (None, False) if no rect fits.
    """
    best_rect = None
    best_rotated = False
    best_score = float('inf')

    orientations = [(pw, ph, False)]
    if allow_rotation and pw != ph:
        orientations.append((ph, pw, True))

    for rect in free_rects:
        for w, h, rotated in orientations:
            if rect.width >= w and rect.height >= h:
                score = min(rect.width - w, rect.height - h)
                if score < best_score:
                    best_score = score
                    best_rect = rect
                    best_rotated = rotated

    return best_rect, best_rotated


def _guillotine_split(free_rects, rect, pw, ph):
    """
    After placing a pw × ph part at rect's origin, split the remaining space.

    Uses Short Axis Split: split along whichever leftover dimension is smaller.
    """
    free_rects.remove(rect)
    right_w = rect.width - pw
    top_h = rect.height - ph

    if right_w < top_h:
        # Horizontal split: right strip at part height; top strip full width
        if right_w > 0:
            free_rects.append(FreeRect(rect.x + pw, rect.y, right_w, ph))
        if top_h > 0:
            free_rects.append(FreeRect(rect.x, rect.y + ph, rect.width, top_h))
    else:
        # Vertical split: top strip at part width; right strip full height
        if top_h > 0:
            free_rects.append(FreeRect(rect.x, rect.y + ph, pw, top_h))
        if right_w > 0:
            free_rects.append(FreeRect(rect.x + pw, rect.y, right_w, rect.height))


def optimize_sheets(parts_in: list, sheet_width: float, sheet_height: float,
                    allow_rotation: bool = True) -> SheetPlan:
    """
    Pack rectangular parts onto sheets using the Guillotine cut algorithm.

    parts_in:       list of (width, height, label) tuples in inches
    sheet_width:    sheet width in inches
    sheet_height:   sheet height in inches
    allow_rotation: if True, parts may be rotated 90 degrees to improve fit

    Returns a SheetPlan. Parts that cannot fit on any single sheet are placed
    in SheetPlan.oversized.
    """
    oversized = []
    valid = []

    for pw, ph, label in parts_in:
        fits = ((pw <= sheet_width and ph <= sheet_height) or
                (allow_rotation and ph <= sheet_width and pw <= sheet_height))
        if fits:
            valid.append((pw, ph, label))
        else:
            oversized.append((pw, ph, label))

    # Largest area first
    valid.sort(key=lambda p: p[0] * p[1], reverse=True)

    sheets = []
    for pw, ph, label in valid:
        placed = False
        for sheet in sheets:
            rect, rotated = _best_fit(sheet._free_rects, pw, ph, allow_rotation)
            if rect is not None:
                w, h = (ph, pw) if rotated else (pw, ph)
                sheet.placed.append(PlacedPart(rect.x, rect.y, w, h, label, rotated))
                _guillotine_split(sheet._free_rects, rect, w, h)
                placed = True
                break
        if not placed:
            new_sheet = Sheet(sheet_width, sheet_height)
            rect, rotated = _best_fit(new_sheet._free_rects, pw, ph, allow_rotation)
            w, h = (ph, pw) if rotated else (pw, ph)
            new_sheet.placed.append(PlacedPart(rect.x, rect.y, w, h, label, rotated))
            _guillotine_split(new_sheet._free_rects, rect, w, h)
            sheets.append(new_sheet)

    return SheetPlan(sheets=sheets, sheet_width=sheet_width,
                     sheet_height=sheet_height, oversized=oversized)


def format_sheet_plan_text(plan: SheetPlan) -> str:
    """Render a SheetPlan as a plain-text sheet optimization section."""
    w, h = plan.sheet_width, plan.sheet_height
    lines = [
        f'Sheet Optimization  (sheet: {w:.0f} x {h:.0f} in)',
        '=' * 44,
    ]
    for i, sheet in enumerate(plan.sheets, 1):
        parts_str = '  '.join(
            f'[{p.width:.1f}x{p.height:.1f}{"R" if p.rotated else ""}]'
            for p in sheet.placed
        )
        lines.append(
            f'Sheet {i}:  {parts_str}  '
            f'waste: {sheet.waste_area:.1f} in\u00b2 ({sheet.waste_pct:.1f}%)'
        )
    lines.append('')
    lines.append(f'Total sheets: {len(plan.sheets)} x {w:.0f}x{h:.0f} in')
    lines.append(
        f'Total waste:  {plan.total_waste_area:.1f} in\u00b2  ({plan.waste_pct:.1f}%)'
    )
    if plan.oversized:
        lines.append('')
        lines.append('Parts too large for sheet size (not included):')
        for pw, ph, label in plan.oversized:
            lines.append(f'  {label}: {pw:.1f} x {ph:.1f} in')
    return '\n'.join(lines)


def format_sheet_plan_html(plan: SheetPlan) -> str:
    """Render SheetPlan statistics as an HTML section (no SVG diagram)."""
    import html as _html
    w, h = plan.sheet_width, plan.sheet_height

    rows = []
    for i, sheet in enumerate(plan.sheets, 1):
        parts_str = ', '.join(
            f'{p.width:.1f}&times;{p.height:.1f}'
            + ('<em>R</em>' if p.rotated else '')
            for p in sheet.placed
        )
        rows.append(
            f'<tr><td>{i}</td><td>{parts_str}</td>'
            f'<td>{sheet.waste_area:.1f} in\u00b2 ({sheet.waste_pct:.1f}%)</td></tr>'
        )

    thead = '<tr><th>Sheet</th><th>Parts</th><th>Waste</th></tr>'
    tbody = ''.join(rows)
    totals = (
        f'<p>Total sheets: {len(plan.sheets)} &times; {w:.0f}&times;{h:.0f} in<br>'
        f'Total waste: {plan.total_waste_area:.1f} in\u00b2 ({plan.waste_pct:.1f}%)</p>'
    )

    oversized_html = ''
    if plan.oversized:
        items_html = ''.join(
            f'<li>{_html.escape(label)}: {pw:.1f} &times; {ph:.1f} in</li>'
            for pw, ph, label in plan.oversized
        )
        oversized_html = (
            f'<p><strong>Warning:</strong> parts too large for sheet size '
            f'(not included):</p><ul>{items_html}</ul>'
        )

    return (
        f'<h2>Sheet Optimization ({w:.0f}&times;{h:.0f} in)</h2>'
        f'<table><thead>{thead}</thead><tbody>{tbody}</tbody></table>'
        f'{totals}{oversized_html}'
    )
