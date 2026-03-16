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

    def format(self, cutlist: CutList):
        from .optimizer import format_plan_html
        items = cutlist.sorted_items()
        title = html.escape(self.docname)
        header = ''.join(f'<th>{html.escape(h)}</th>' for h in self.fieldnames)
        rows = ''.join(self.item_to_row(item) for item in items)
        summary_html = self._material_summary_html(items)

        plan = self._run_optimizer(items)
        cutplan_html = format_plan_html(plan) if plan is not None else ''

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
