"""
Cut list optimizer using First Fit Decreasing (FFD) bin packing.

All lengths are in inches. This module has no Fusion 360 API dependencies
and can be unit-tested standalone.
"""

from dataclasses import dataclass, field

DEFAULT_KERF = 0.125          # 1/8" saw blade kerf
DEFAULT_MIN_OFFCUT = 12.0     # 12" minimum useful off-cut

# Parts whose height (smallest dimension) is at or below this threshold are
# treated as sheet goods (plywood/MDF) and excluded from the optimizer.
SHEET_GOODS_MAX_HEIGHT_IN = 0.75


@dataclass
class Board:
    """A single board with the parts assigned to it."""

    stock_length: float
    kerf: float
    parts: list = field(default_factory=list)

    @property
    def used(self) -> float:
        """Material consumed: each placed part costs part_length + kerf."""
        return sum(p + self.kerf for p in self.parts)

    @property
    def remaining(self) -> float:
        return self.stock_length - self.used

    @property
    def waste(self) -> float:
        return self.remaining


@dataclass
class Plan:
    """Result of a cut optimization run."""

    boards: list
    stock_length: float
    kerf: float
    min_offcut: float
    sheet_goods_skipped: bool = False

    @property
    def total_waste(self) -> float:
        return sum(b.waste for b in self.boards)

    @property
    def total_length(self) -> float:
        return len(self.boards) * self.stock_length

    @property
    def waste_pct(self) -> float:
        if self.total_length == 0:
            return 0.0
        return (self.total_waste / self.total_length) * 100

    @property
    def offcuts(self) -> list:
        """Waste pieces long enough to be reusable (>= min_offcut)."""
        return [b.waste for b in self.boards if b.waste >= self.min_offcut]


def parse_stock_lengths(value: str) -> list:
    """
    Parse a comma-separated string of stock lengths into a list of floats.

    Returns an empty list if the string is blank or contains no valid numbers.
    """
    lengths = []
    for token in value.split(','):
        token = token.strip()
        if token:
            try:
                lengths.append(float(token))
            except ValueError:
                pass
    return sorted(lengths, reverse=True)


def optimize(parts_in: list, stock_lengths: list,
             kerf: float = DEFAULT_KERF,
             min_offcut: float = DEFAULT_MIN_OFFCUT) -> Plan:
    """
    Assign parts to boards using First Fit Decreasing bin packing.

    parts_in:      flat list of part lengths in inches (already expanded by count)
    stock_lengths: available stock lengths in inches (the longest is used)
    kerf:          saw blade kerf deducted per placement, in inches
    min_offcut:    minimum off-cut length to treat as usable, in inches

    Returns a Plan describing the board assignments.
    """
    stock_length = max(stock_lengths)
    parts = sorted(parts_in, reverse=True)
    boards = []

    for part in parts:
        placed = False
        for board in boards:
            if board.remaining >= part + kerf:
                board.parts.append(part)
                placed = True
                break
        if not placed:
            new_board = Board(stock_length=stock_length, kerf=kerf)
            new_board.parts.append(part)
            boards.append(new_board)

    return Plan(boards=boards, stock_length=stock_length,
                kerf=kerf, min_offcut=min_offcut)


def format_plan_text(plan: Plan) -> str:
    """Render a Plan as a plain-text cut optimization section."""
    lines = [
        f'Cut Optimization  (stock: {plan.stock_length:.0f} in)',
        '=' * 44,
    ]

    for i, board in enumerate(plan.boards, 1):
        parts_str = ''.join(f'[{p:.1f}]' for p in board.parts)
        lines.append(f'Board {i}:  {parts_str}  waste: {board.waste:.1f} in')

    lines.append('')
    lines.append(f'Total boards: {len(plan.boards)} x {plan.stock_length:.0f} in')
    lines.append(f'Total waste:  {plan.total_waste:.1f} in  ({plan.waste_pct:.1f}%)')

    offcuts = plan.offcuts
    if offcuts:
        offcut_str = ' '.join(f'[{o:.1f}]' for o in offcuts)
        lines.append(
            f'Off-cuts (>= {plan.min_offcut:.0f} in):  {offcut_str}'
            f'  \u2014 usable for other projects'
        )

    if plan.sheet_goods_skipped:
        lines.append('')
        lines.append('Note: sheet goods excluded from cut optimization.')

    return '\n'.join(lines)


def format_plan_html(plan: Plan) -> str:
    """Render a Plan as an HTML cut optimization section."""
    import html as _html

    board_rows = []
    for i, board in enumerate(plan.boards, 1):
        parts_str = _html.escape(''.join(f'[{p:.1f}]' for p in board.parts))
        board_rows.append(
            f'<tr><td>{i}</td><td>{parts_str}</td><td>{board.waste:.1f} in</td></tr>'
        )

    thead = '<tr><th>Board</th><th>Parts</th><th>Waste</th></tr>'
    tbody = ''.join(board_rows)

    totals = (
        f'<p>Total boards: {len(plan.boards)} &times; {plan.stock_length:.0f} in<br>'
        f'Total waste: {plan.total_waste:.1f} in ({plan.waste_pct:.1f}%)</p>'
    )

    offcut_html = ''
    offcuts = plan.offcuts
    if offcuts:
        offcut_str = _html.escape(' '.join(f'[{o:.1f}]' for o in offcuts))
        offcut_html = (
            f'<p><strong>Off-cuts (&ge; {plan.min_offcut:.0f} in):</strong>'
            f' {offcut_str} &mdash; usable for other projects</p>'
        )

    sheet_html = ''
    if plan.sheet_goods_skipped:
        sheet_html = '<p><em>Note: sheet goods excluded from cut optimization.</em></p>'

    return (
        f'<h2>Cut Optimization (stock: {plan.stock_length:.0f} in)</h2>'
        f'<table><thead>{thead}</thead><tbody>{tbody}</tbody></table>'
        f'{totals}{offcut_html}{sheet_html}'
    )


def format_plan_csv(plan: Plan) -> str:
    """Render a Plan as a CSV string (for the _cutplan file)."""
    import csv
    import io

    with io.StringIO(newline='') as f:
        w = csv.writer(f, dialect='excel')
        w.writerow(['board', 'parts', 'waste_in'])
        for i, board in enumerate(plan.boards, 1):
            parts_str = ','.join(f'{p:.1f}' for p in board.parts)
            w.writerow([i, parts_str, f'{board.waste:.1f}'])

        f.write('\n')
        w.writerow(['total_boards', 'stock_length_in', 'total_waste_in', 'waste_pct',
                    'kerf_in', 'min_offcut_in'])
        w.writerow([
            len(plan.boards),
            f'{plan.stock_length:.1f}',
            f'{plan.total_waste:.1f}',
            f'{plan.waste_pct:.1f}',
            f'{plan.kerf:.4f}',
            f'{plan.min_offcut:.1f}',
        ])

        return f.getvalue()


def format_plan_json(plan: Plan) -> str:
    """Render a Plan as a JSON string (for the _cutplan file)."""
    import json

    boards = [
        {
            'board': i,
            'parts_in': [round(p, 4) for p in board.parts],
            'waste_in': round(board.waste, 4),
        }
        for i, board in enumerate(plan.boards, 1)
    ]

    return json.dumps({
        'stock_length_in': plan.stock_length,
        'kerf_in': plan.kerf,
        'min_offcut_in': plan.min_offcut,
        'boards': boards,
        'totals': {
            'board_count': len(plan.boards),
            'total_length_in': round(plan.total_length, 4),
            'total_waste_in': round(plan.total_waste, 4),
            'waste_pct': round(plan.waste_pct, 2),
        },
        'offcuts_in': [round(o, 4) for o in plan.offcuts],
        'sheet_goods_skipped': plan.sheet_goods_skipped,
    }, indent=2)
