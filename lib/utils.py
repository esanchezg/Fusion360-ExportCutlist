"""Shared utilities for board-feet calculation and material summary formatting."""

from collections import defaultdict

# Fusion 360 stores dimensions in centimeters internally.
CM_TO_IN = 0.393701

# Default waste factor for buy-quantity calculation (10%).
WASTE_FACTOR = 0.10

# Unit strings that indicate imperial measurement.
_IMPERIAL_UNITS = {'in', 'ft'}


def is_imperial(units: str) -> bool:
    """Return True if the unit string indicates imperial measurement."""
    return units in _IMPERIAL_UNITS


def board_feet(length_cm: float, width_cm: float, height_cm: float) -> float:
    """Return board feet for a single part whose dimensions are in centimeters."""
    l = length_cm * CM_TO_IN
    w = width_cm * CM_TO_IN
    h = height_cm * CM_TO_IN
    return (l * w * h) / 144.0


def volume_cm3(length_cm: float, width_cm: float, height_cm: float) -> float:
    """Return volume in cubic centimeters."""
    return length_cm * width_cm * height_cm


def material_summary(items, units: str, waste_factor: float = WASTE_FACTOR) -> list:
    """
    Compute per-material totals from a list of CutListItems.

    Returns a list of dicts sorted by material name. Each dict has:
      material, count, value, unit_label
    Imperial units also include: buy (value with waste factor applied).
    """
    imperial = is_imperial(units)
    totals = defaultdict(lambda: {'count': 0, 'value': 0.0})

    for item in items:
        dims = item.dimensions
        if imperial:
            val = board_feet(dims.length, dims.width, dims.height) * item.count
        else:
            val = volume_cm3(dims.length, dims.width, dims.height) * item.count
        totals[item.material]['count'] += item.count
        totals[item.material]['value'] += val

    result = []
    for mat in sorted(totals):
        entry = {
            'material': mat,
            'count': totals[mat]['count'],
            'value': totals[mat]['value'],
            'unit_label': 'bf' if imperial else 'cm\u00b3',
        }
        if imperial:
            entry['buy'] = totals[mat]['value'] * (1 + waste_factor)
        result.append(entry)

    return result


def format_material_summary(items, units: str, waste_factor: float = WASTE_FACTOR) -> str:
    """
    Produce a plain-text Material Summary block from a list of CutListItems.

    Imperial designs show board feet and a buy quantity with the waste factor.
    Metric designs show volume in cm³ and omit the waste/buy line.
    """
    rows = material_summary(items, units, waste_factor)
    if not rows:
        return ''

    imperial = is_imperial(units)
    pct = int(waste_factor * 100)
    lines = ['Material Summary', '=' * 44]

    for row in rows:
        mat = row['material']
        count = row['count']
        val = row['value']
        if imperial:
            buy = row['buy']
            lines.append(
                f'{mat:<24} {count:>4} pcs  {val:>8.2f} bf'
                f'    (buy ~{buy:.2f} bf with {pct}% waste)'
            )
        else:
            lines.append(f'{mat:<24} {count:>4} pcs  {val:>10.1f} cm\u00b3')

    return '\n'.join(lines)
