import csv
import html
import io
import json
import re
import textwrap
import typing

from dataclasses import dataclass

import adsk.core

from .texttable import Texttable
from .cutlist import CutList, CutListItem
from .utils import (
    CM_TO_IN, WASTE_FACTOR, is_imperial, board_feet, volume_cm3,
    material_summary, format_material_summary,
)


@dataclass
class FormatOptions:
    """
    Options that affect cutlist formatting.

    Options:
      component_names          use component names instead of body names
      short_names              only use the final body or component name, rather than the full path
      remove_numeric_suffixes  remove common numeric suffixes from names
      unique_names             only output unique names for each item
      include_material         include material in the output if the format supports it
      name_separator           the separator to use when joining name elements
      units                    the units to use for dimensions
      stock_lengths            comma-separated stock lengths in inches; empty = skip optimizer
      kerf_in                  saw blade kerf in inches
      min_offcut_in            minimum off-cut length in inches to treat as usable
    """
    component_names: bool = False
    short_names: bool = False
    remove_numeric_suffixes: bool = False
    unique_names: bool = False
    include_material: bool = True
    name_separator: str = '/'
    units: str = 'auto'
    stock_lengths: str = ''
    kerf_in: float = 0.125
    min_offcut_in: float = 12.0
    sheet_size: str = ''
    respect_grain: bool = False


class FileFilter:
    def __init__(self, name, ext):
        self.name = name
        self.ext = ext

    @property
    def filter_str(self):
        return f'{self.name} (*.{self.ext})'


class Format:
    name = 'Base Format'
    filefilter = FileFilter('Text Files', 'txt')

    def __init__(self, units_manager: adsk.core.UnitsManager, docname: str, options: FormatOptions):
        self.units_manager = units_manager
        self.docname = docname
        self.options = options
        self.units = options.units if options.units != 'auto' else units_manager.defaultLengthUnits

    @property
    def filename(self) -> str:
        name = self.docname.lower().replace(' ', '_')
        return f'{name}.{self.filefilter.ext}'

    def format_value(self, value, showunits=False) -> str:
        return self.units_manager.formatInternalValue(value, self.units, showunits)

    def format_item_names(self, item: CutListItem) -> list[str]:
        separator = self.options.name_separator

        names = []
        for p in item.paths:
            if self.options.component_names and p.parent_name:
                if self.options.short_names:
                    name = p.parent_name
                else:
                    name = separator.join(p.components)
            else:
                if self.options.short_names:
                    name = p.body_name
                else:
                    name = separator.join((*p.components, p.body_name))

            if self.options.remove_numeric_suffixes:
                name = re.sub(r'(\s+\d+|\s*\(\d+\))$', '', name)

            names.append(name)

        if self.options.unique_names:
            names = set(names)

        return sorted(names)

    def _run_optimizer(self, items):
        """
        Expand items into a flat list of lumber lengths (inches) and run FFD.

        Sheet goods (height <= 0.75 in) are skipped. Returns a Plan, or None
        if stock_lengths is not configured or no lumber parts exist.
        """
        from .optimizer import (parse_stock_lengths, optimize,
                                 SHEET_GOODS_MAX_HEIGHT_IN)
        stock_lengths = parse_stock_lengths(self.options.stock_lengths)
        if not stock_lengths:
            return None

        parts = []
        skipped = False
        for item in items:
            if item.dimensions.height * CM_TO_IN <= SHEET_GOODS_MAX_HEIGHT_IN:
                skipped = True
                continue
            length_in = item.dimensions.length * CM_TO_IN
            parts.extend([length_in] * item.count)

        if not parts:
            return None

        plan = optimize(parts, stock_lengths,
                        self.options.kerf_in, self.options.min_offcut_in)
        plan.sheet_goods_skipped = skipped
        return plan

    def _run_sheet_optimizer(self, items):
        """
        Filter sheet goods from items and run the 2D sheet optimizer.

        Returns a SheetPlan, or None if sheet_size is not configured or no
        sheet goods exist in the item list.
        """
        from .sheet_optimizer import parse_sheet_size, optimize_sheets
        from .optimizer import SHEET_GOODS_MAX_HEIGHT_IN
        sheet_size = parse_sheet_size(self.options.sheet_size)
        if sheet_size is None:
            return None
        sheet_width, sheet_height = sheet_size
        parts = []
        for item in items:
            if item.dimensions.height * CM_TO_IN > SHEET_GOODS_MAX_HEIGHT_IN:
                continue
            l_in = item.dimensions.length * CM_TO_IN
            w_in = item.dimensions.width * CM_TO_IN
            names = self.format_item_names(item)
            label = names[0] if names else f'{l_in:.1f}x{w_in:.1f}'
            parts.extend([(l_in, w_in, label)] * item.count)
        if not parts:
            return None
        allow_rotation = not self.options.respect_grain
        return optimize_sheets(parts, sheet_width, sheet_height,
                               allow_rotation=allow_rotation)

    def format_cutplan(self, cutlist: CutList):
        """
        Return a separate cut-plan string for formats that write a second file.

        Returns None by default; CSV and JSON override this.
        """
        return None

    def format(self, cutlist: CutList):
        raise NotImplementedError


class JSONFormat(Format):
    name = 'JSON'
    filefilter = FileFilter('JSON Files', 'json')

    def item_to_dict(self, item: CutListItem):
        include_material = self.options.include_material
        dims = item.dimensions
        if is_imperial(self.units):
            bf_entry = {'board_feet': round(board_feet(dims.length, dims.width, dims.height), 4)}
        else:
            bf_entry = {'volume_cm3': round(volume_cm3(dims.length, dims.width, dims.height), 4)}
        return {
            'count': item.count,
            'dimensions': {
                'units': self.units,
                'length': self.format_value(dims.length),
                'width': self.format_value(dims.width),
                'height': self.format_value(dims.height),
            },
            **bf_entry,
            **({'material': item.material} if include_material else {}),
            'names': self.format_item_names(item),
        }

    def format(self, cutlist: CutList):
        items = cutlist.sorted_items()
        rows = material_summary(items, self.units)
        summary = []
        for row in rows:
            entry = {
                'material': row['material'],
                'count': row['count'],
                'unit_label': row['unit_label'],
                'value': round(row['value'], 4),
            }
            if is_imperial(self.units):
                entry['buy'] = round(row['buy'], 4)
            summary.append(entry)
        return json.dumps({
            'items': [self.item_to_dict(item) for item in items],
            'material_summary': summary,
        }, indent=2)

    def format_cutplan(self, cutlist: CutList):
        """Return a JSON cut-plan string, or None if the optimizer is not configured."""
        from .optimizer import format_plan_json
        plan = self._run_optimizer(cutlist.sorted_items())
        if plan is None:
            return None
        return format_plan_json(plan)


class CSVDictBuilder:
    def __init__(self, fields: list[str]):
        self.fields = fields
        self.index = 0
        self.dict = {}

    def set_field(self, value: str):
        if self.index < len(self.fields):
            field = self.fields[self.index]
            self.dict[field] = value
            self.index += 1

    def build(self) -> dict:
        return self.dict


class CSVFormat(Format):
    name = 'CSV'
    filefilter = FileFilter('CSV Files', 'csv')

    dialect = 'excel'
    include_board_feet = True

    @property
    def fieldnames(self):
        include_material = self.options.include_material
        lengthkey, widthkey, heightkey = [f'{v} ({self.units})' for v in ['length', 'width', 'height']]
        bf_key = 'board_ft' if is_imperial(self.units) else 'vol_cm3'
        return [
            'count',
            *(['material'] if include_material else []),
            lengthkey,
            widthkey,
            heightkey,
            bf_key,
            'names',
        ]

    def item_to_dict(self, item: CutListItem):
        d = CSVDictBuilder(self.fieldnames)
        dims = item.dimensions

        d.set_field(item.count)

        if self.options.include_material:
            d.set_field(item.material)

        d.set_field(self.format_value(dims.length))
        d.set_field(self.format_value(dims.width))
        d.set_field(self.format_value(dims.height))

        if is_imperial(self.units):
            d.set_field(f'{board_feet(dims.length, dims.width, dims.height):.2f}')
        else:
            d.set_field(f'{volume_cm3(dims.length, dims.width, dims.height):.1f}')

        d.set_field(','.join(self.format_item_names(item)))

        return d.build()

    def format(self, cutlist: CutList):
        items = cutlist.sorted_items()
        with io.StringIO(newline='') as f:
            w = csv.DictWriter(f, dialect=self.dialect, fieldnames=self.fieldnames)
            w.writeheader()
            w.writerows([self.item_to_dict(item) for item in items])

            if self.include_board_feet:
                rows = material_summary(items, self.units)
                if rows:
                    f.write('\n')
                    imperial = is_imperial(self.units)
                    summary_fields = (['material', 'count', 'board_ft', 'buy_bf']
                                      if imperial else ['material', 'count', 'vol_cm3'])
                    sw = csv.DictWriter(f, dialect=self.dialect, fieldnames=summary_fields)
                    sw.writeheader()
                    for row in rows:
                        r = {'material': row['material'], 'count': row['count']}
                        if imperial:
                            r['board_ft'] = f"{row['value']:.2f}"
                            r['buy_bf'] = f"{row['buy']:.2f}"
                        else:
                            r['vol_cm3'] = f"{row['value']:.1f}"
                        sw.writerow(r)

            return f.getvalue()

    def format_cutplan(self, cutlist: CutList):
        """Return a CSV cut-plan string, or None if the optimizer is not configured."""
        from .optimizer import format_plan_csv
        plan = self._run_optimizer(cutlist.sorted_items())
        if plan is None:
            return None
        return format_plan_csv(plan)


class CutlistOptimizerFormat(CSVFormat):
    '''
    CSV format used by https://cutlistoptimizer.com/
    '''

    name = 'Cutlist Optimizer'
    include_board_feet = False

    @property
    def fieldnames(self):
        include_material = self.options.include_material
        return [
            'Length',
            'Width',
            'Qty',
            *(['Material'] if include_material else []),
            'Label',
            'Enabled',
        ]

    def item_to_dict(self, item: CutListItem):
        # CutlistOptimizer uses str.split to 'parse' the fields in each record.
        # Import will fail when fields contain the delimiter. Use semicolon to
        # separate the names and remove all delimiters from str values.

        d = CSVDictBuilder(self.fieldnames)

        d.set_field(self.format_value(item.dimensions.length))
        d.set_field(self.format_value(item.dimensions.width))
        d.set_field(item.count)

        if self.options.include_material:
            d.set_field(item.material.replace(',', ''))

        d.set_field(';'.join(n.replace(',', '') for n in self.format_item_names(item)))

        d.set_field('true')

        return d.build()


class CutlistEvoFormat(CSVFormat):
    '''
    Tab-separated format used by https://cutlistevo.com/
    '''

    name = 'Cutlist Evo'
    include_board_feet = False
    filefilter = FileFilter('Text Files', 'txt')

    dialect = 'excel-tab'

    @property
    def fieldnames(self):
        include_material = self.options.include_material
        return [
            'Length',
            'Width',
            'Thickness',
            'Quantity',
            'Rotation',
            'Name',
            *(['Material'] if include_material else []),
            'Banding'
        ]

    def item_to_dict(self, item: CutListItem):
        d = CSVDictBuilder(self.fieldnames)

        d.set_field(self.format_value(item.dimensions.length))
        d.set_field(self.format_value(item.dimensions.width))
        d.set_field(self.format_value(item.dimensions.height))
        d.set_field(item.count)
        d.set_field(','.join(['L'] * item.count))
        d.set_field(','.join(self.format_item_names(item)))

        if self.options.include_material:
            d.set_field(item.material)

        d.set_field(','.join(['N'] * item.count))

        return d.build()


class TableFormat(Format):
    name = 'Table'

    @property
    def fieldnames(self):
        include_material = self.options.include_material
        lengthkey, widthkey, heightkey = [f'{v} ({self.units})' for v in ['length', 'width', 'height']]
        bf_key = 'board ft' if is_imperial(self.units) else 'vol (cm\u00b3)'
        return [
            'count',
            *(['material'] if include_material else []),
            lengthkey,
            widthkey,
            heightkey,
            bf_key,
            'names',
        ]

    def item_to_row(self, item: CutListItem):
        include_material = self.options.include_material
        dims = item.dimensions
        if is_imperial(self.units):
            bf_val = f'{board_feet(dims.length, dims.width, dims.height):.2f}'
        else:
            bf_val = f'{volume_cm3(dims.length, dims.width, dims.height):.1f}'
        return [
            item.count,
            *([item.material] if include_material else []),
            self.format_value(dims.length),
            self.format_value(dims.width),
            self.format_value(dims.height),
            bf_val,
            '\n'.join(self.format_item_names(item)),
        ]

    def format(self, cutlist: CutList):
        from .optimizer import format_plan_text
        include_material = self.options.include_material
        items = cutlist.sorted_items()

        tt = Texttable(max_width=0)
        tt.set_deco(Texttable.HEADER | Texttable.HLINES)
        tt.header(self.fieldnames)
        tt.set_cols_dtype(['i', *(['t'] if include_material else []), 't', 't', 't', 't', 't'])
        tt.set_cols_align(['r', *(['l'] if include_material else []), 'r', 'r', 'r', 'r', 'l'])
        tt.add_rows([self.item_to_row(item) for item in items], header=False)

        output = tt.draw()
        output += '\n\n' + format_material_summary(items, self.units)

        plan = self._run_optimizer(items)
        if plan is not None:
            output += '\n\n' + format_plan_text(plan)

        from .sheet_optimizer import format_sheet_plan_text
        sheet_plan = self._run_sheet_optimizer(items)
        if sheet_plan is not None:
            output += '\n\n' + format_sheet_plan_text(sheet_plan)

        return output


class HTMLFormat(Format):
    name = 'HTML'
    filefilter = FileFilter('HTML Files', 'html')

    @property
    def fieldnames(self):
        include_material = self.options.include_material
        lengthkey, widthkey, heightkey = [f'{v} ({self.units})' for v in ['Length', 'Width', 'Height']]
        bf_key = 'Board ft' if is_imperial(self.units) else 'Vol (cm\u00b3)'
        return [
            'Count',
            lengthkey,
            widthkey,
            heightkey,
            bf_key,
            *(['Material'] if include_material else []),
            'Names',
        ]

    def item_to_row(self, item: CutListItem):
        include_material = self.options.include_material
        dims = item.dimensions
        if is_imperial(self.units):
            bf_val = f'{board_feet(dims.length, dims.width, dims.height):.2f}'
        else:
            bf_val = f'{volume_cm3(dims.length, dims.width, dims.height):.1f}'
        cols = [
            item.count,
            self.format_value(dims.length),
            self.format_value(dims.width),
            self.format_value(dims.height),
            bf_val,
            *([html.escape(item.material)] if include_material else []),
            '<br>'.join(html.escape(n) for n in self.format_item_names(item)),
        ]
        return '<tr>' + ''.join(f'<td>{c}</td>' for c in cols) + '</tr>'

    def _material_summary_html(self, items) -> str:
        """Render the material summary as an HTML section."""
        rows = material_summary(items, self.units)
        if not rows:
            return ''

        imperial = is_imperial(self.units)
        pct = int(WASTE_FACTOR * 100)

        if imperial:
            header = '<tr><th>Material</th><th>Pcs</th><th>Board ft</th><th>Buy (with waste)</th></tr>'
        else:
            header = '<tr><th>Material</th><th>Pcs</th><th>Vol (cm\u00b3)</th></tr>'

        tr_rows = []
        for row in rows:
            mat = html.escape(row['material'])
            val = f"{row['value']:.2f}" if imperial else f"{row['value']:.1f}"
            if imperial:
                buy = f"~{row['buy']:.2f} bf ({pct}% waste)"
                tr_rows.append(f'<tr><td>{mat}</td><td>{row["count"]}</td><td>{val}</td><td>{buy}</td></tr>')
            else:
                tr_rows.append(f'<tr><td>{mat}</td><td>{row["count"]}</td><td>{val}</td></tr>')

        trs = ''.join(tr_rows)
        return f'<h2>Material Summary</h2><table><thead>{header}</thead><tbody>{trs}</tbody></table>'

    def _cut_visualization_html(self, plan, items) -> str:
        """Render an SVG cut diagram for a lumber Plan."""
        from .optimizer import SHEET_GOODS_MAX_HEIGHT_IN
        # Assign a consistent color to each unique part length so the same
        # dimension is always the same color across all boards.
        PALETTE = [
            '#5b9bd5', '#ed7d31', '#70ad47', '#ffc000',
            '#4472c4', '#c55a11', '#548235', '#7030a0',
        ]
        unique_lengths = sorted(
            {p for board in plan.boards for p in board.parts}, reverse=True)
        color_map = {l: PALETTE[i % len(PALETTE)]
                     for i, l in enumerate(unique_lengths)}

        # Map rounded length → part name for tooltips and legend
        length_to_name = {}
        for item in items:
            if item.dimensions.height * CM_TO_IN <= SHEET_GOODS_MAX_HEIGHT_IN:
                continue
            l_in = round(item.dimensions.length * CM_TO_IN, 1)
            names = self.format_item_names(item)
            if names:
                length_to_name[l_in] = names[0]

        BAR_H    = 34   # height of each colored bar
        ROW_H    = 48   # row height including vertical gap
        LABEL_W  = 72   # pixels reserved for the "Board N" label
        CONTENT_W = 740 # pixels for the actual bar area
        RIGHT_PAD = 20
        SVG_W = LABEL_W + CONTENT_W + RIGHT_PAD
        SVG_H = len(plan.boards) * ROW_H + 16

        elems = []
        for i, board in enumerate(plan.boards):
            bar_y = i * ROW_H + 7

            # Board label
            elems.append(
                f'<text x="{LABEL_W - 6}" y="{bar_y + BAR_H // 2 + 4}" '
                f'text-anchor="end" font-size="12" fill="#444">'
                f'Board {i + 1}</text>'
            )

            x = float(LABEL_W)
            for part in board.parts:
                # Each part consumes part_length + kerf of stock; using that
                # here ensures the bar fills exactly to the stock length.
                seg_w = ((part + plan.kerf) / plan.stock_length) * CONTENT_W
                color = color_map[part]
                name = length_to_name.get(round(part, 1), '')
                tip = html.escape(
                    f'{name}: {part:.1f} in' if name else f'{part:.1f} in'
                )
                elems.append(
                    f'<rect x="{x:.2f}" y="{bar_y}" '
                    f'width="{max(seg_w, 1):.2f}" height="{BAR_H}" '
                    f'fill="{color}" stroke="#fff" stroke-width="1.5">'
                    f'<title>{tip}</title></rect>'
                )
                if seg_w >= 36:
                    elems.append(
                        f'<text x="{x + seg_w / 2:.2f}" '
                        f'y="{bar_y + BAR_H // 2 + 4}" '
                        f'text-anchor="middle" font-size="11" '
                        f'fill="#fff" font-weight="bold">'
                        f'{part:.1f}&quot;</text>'
                    )
                x += seg_w

            # Waste segment
            waste_w = (board.waste / plan.stock_length) * CONTENT_W
            if waste_w >= 1:
                tip = html.escape(f'waste: {board.waste:.1f} in')
                elems.append(
                    f'<rect x="{x:.2f}" y="{bar_y}" '
                    f'width="{waste_w:.2f}" height="{BAR_H}" '
                    f'fill="#e0e0e0" stroke="#fff" stroke-width="1.5">'
                    f'<title>{tip}</title></rect>'
                )
                if waste_w >= 32:
                    elems.append(
                        f'<text x="{x + waste_w / 2:.2f}" '
                        f'y="{bar_y + BAR_H // 2 + 4}" '
                        f'text-anchor="middle" font-size="10" fill="#999">'
                        f'{board.waste:.1f}&quot;</text>'
                    )

        svg = (
            f'<svg width="{SVG_W}" height="{SVG_H}" '
            f'font-family="sans-serif" '
            f'xmlns="http://www.w3.org/2000/svg">\n  '
            + '\n  '.join(elems)
            + '\n</svg>'
        )

        # Legend: one swatch per unique part length + waste
        swatch = ('display:inline-block;width:14px;height:14px;'
                  'border-radius:2px;margin-right:4px;vertical-align:middle;')
        item_style = 'display:inline-flex;align-items:center;margin-right:14px;font-size:12px;'
        legend_parts = []
        for l in sorted(color_map, reverse=True):
            name = length_to_name.get(round(l, 1), '')
            label = (f'{html.escape(name)} ({l:.1f}&quot;)' if name
                     else f'{l:.1f}&quot;')
            legend_parts.append(
                f'<span style="{item_style}">'
                f'<span style="{swatch}background:{color_map[l]};"></span>'
                f'{label}</span>'
            )
        legend_parts.append(
            f'<span style="{item_style}">'
            f'<span style="{swatch}background:#e0e0e0;"></span>'
            f'waste</span>'
        )
        legend = f'<p style="margin-top:6px;">{"".join(legend_parts)}</p>'

        sheet_note = ('<p><em>Sheet goods are not included in the cut diagram.</em></p>'
                      if plan.sheet_goods_skipped else '')

        return f'<h2>Cut Diagram</h2>{svg}{legend}{sheet_note}'

    def _sheet_visualization_html(self, plan) -> str:
        """Render SVG cut diagrams, one per sheet, for a SheetPlan."""
        PALETTE = [
            '#5b9bd5', '#ed7d31', '#70ad47', '#ffc000',
            '#4472c4', '#c55a11', '#548235', '#7030a0',
            '#2e75b6', '#c00000', '#00b050', '#7f7f7f',
        ]
        unique_labels = sorted({p.label for s in plan.sheets for p in s.placed})
        color_map = {l: PALETTE[i % len(PALETTE)] for i, l in enumerate(unique_labels)}

        # Scale to fit within display bounds while preserving aspect ratio
        DISP_W, DISP_H = 400, 600
        scale = min(DISP_W / plan.sheet_width, DISP_H / plan.sheet_height)
        svg_w = int(plan.sheet_width * scale)
        svg_h = int(plan.sheet_height * scale)

        sheet_divs = []
        for i, sheet in enumerate(plan.sheets, 1):
            elems = []
            elems.append(
                f'<rect width="{svg_w}" height="{svg_h}" '
                f'fill="#f8f8f8" stroke="#555" stroke-width="1.5"/>'
            )
            for part in sheet.placed:
                px = part.x * scale
                py = part.y * scale
                pw = part.width * scale
                ph = part.height * scale
                color = color_map.get(part.label, '#ccc')
                rot_note = ' (rotated)' if part.rotated else ''
                tip = html.escape(
                    f'{part.label}: {part.width:.1f}×{part.height:.1f} in{rot_note}'
                )
                elems.append(
                    f'<rect x="{px:.2f}" y="{py:.2f}" '
                    f'width="{pw:.2f}" height="{ph:.2f}" '
                    f'fill="{color}" fill-opacity="0.82" '
                    f'stroke="#fff" stroke-width="1">'
                    f'<title>{tip}</title></rect>'
                )
                if pw >= 40 and ph >= 16:
                    # Show name on first line, dimensions on second if room
                    label_y = py + ph / 2
                    if ph >= 30:
                        label_y = py + ph / 2 - 6
                        elems.append(
                            f'<text x="{px + pw / 2:.2f}" y="{label_y + 14:.2f}" '
                            f'text-anchor="middle" dominant-baseline="middle" '
                            f'font-size="9" fill="#fff" fill-opacity="0.8">'
                            f'{part.width:.1f}&times;{part.height:.1f}</text>'
                        )
                    elems.append(
                        f'<text x="{px + pw / 2:.2f}" y="{label_y:.2f}" '
                        f'text-anchor="middle" dominant-baseline="middle" '
                        f'font-size="10" fill="#fff" font-weight="bold">'
                        f'{html.escape(part.label)}</text>'
                    )

            svg = (
                f'<svg width="{svg_w}" height="{svg_h}" font-family="sans-serif" '
                f'xmlns="http://www.w3.org/2000/svg">'
                + ''.join(elems)
                + '</svg>'
            )
            title = f'Sheet {i} &mdash; {sheet.waste_pct:.1f}% waste'
            sheet_divs.append(
                f'<div style="display:inline-block;margin:8px;'
                f'vertical-align:top;text-align:center;">'
                f'<p style="font-size:12px;margin:0 0 4px;">{title}</p>'
                f'{svg}</div>'
            )

        swatch = ('display:inline-block;width:14px;height:14px;'
                  'border-radius:2px;margin-right:4px;vertical-align:middle;')
        item_style = 'display:inline-flex;align-items:center;margin-right:14px;font-size:12px;'
        legend_parts = [
            f'<span style="{item_style}">'
            f'<span style="{swatch}background:{color_map[l]};"></span>'
            f'{html.escape(l)} in</span>'
            for l in sorted(color_map)
        ]
        legend = f'<p style="margin-top:8px;">{"".join(legend_parts)}</p>'

        return (
            f'<h2>Sheet Cut Diagram</h2>'
            f'<div>{"".join(sheet_divs)}</div>'
            f'{legend}'
        )

    def format(self, cutlist: CutList):
        from .optimizer import format_plan_html
        from .sheet_optimizer import format_sheet_plan_html
        items = cutlist.sorted_items()
        title = html.escape(self.docname)
        header = ''.join(f'<th>{html.escape(h)}</th>' for h in self.fieldnames)
        rows = ''.join(self.item_to_row(item) for item in items)
        summary_html = self._material_summary_html(items)

        plan = self._run_optimizer(items)
        cutplan_html = format_plan_html(plan) if plan is not None else ''
        diagram_html = self._cut_visualization_html(plan, items) if plan is not None else ''

        sheet_plan = self._run_sheet_optimizer(items)
        sheet_stats_html = format_sheet_plan_html(sheet_plan) if sheet_plan is not None else ''
        sheet_viz_html = self._sheet_visualization_html(sheet_plan) if sheet_plan is not None else ''

        return textwrap.dedent(f'''\
            <html>
                <title>{title} Cutlist</title>
                <style>
                    table {{
                        border: 1px solid #000;
                        border-collapse: collapse;
                    }}
                    td, th {{
                        border: 1px solid #000;
                        padding: 0.25em 0.5em;
                    }}
                    thead {{
                        background-color: #eee;
                    }}
                </style>
            </html>
            </body>
                <h1>{title} Cutlist</h1>
                <table>
                    <thead>{header}</thead>
                    <tbody>{rows}</tbody>
                </table>
                {summary_html}
                {cutplan_html}
                {diagram_html}
                {sheet_stats_html}
                {sheet_viz_html}
            </body>
        ''')



ALL_FORMATS = [
    TableFormat,
    CSVFormat,
    JSONFormat,
    HTMLFormat,
    CutlistOptimizerFormat,
    CutlistEvoFormat,
]


def get_format(name: str) -> typing.Type[Format]:
    for fmt in ALL_FORMATS:
        if fmt.name == name:
            return fmt
    raise ValueError(f'unknown format: {name}')
