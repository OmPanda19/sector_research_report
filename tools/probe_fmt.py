"""For the new-table regions, report which empty cells already carry a number format / fill."""
import openpyxl
from openpyxl.utils import get_column_letter

LEGACY = {
    'Control Panel': 35, 'Macroeconomic Model': 99, 'Steel Demand Model': 112, 'Steel Supply Model': 88,
    'Capacity Forecast': 94, 'Capacity Expansion Tracker': 75, 'Capacity Utilisation': 64,
    'Steel Price Forecast': 95, 'Raw Material Forecast': 83, 'Cost Curve': 76, 'Revenue Forecast': 93,
    'EBITDA Model': 98, 'Margin Analysis': 87, 'Working Capital Model': 69, 'Cash Flow Model': 95,
    'Capital Allocation': 76, 'Industry Cycle Model': 86, 'Trade Model': 73, 'ESG Model': 94,
    'Comparable Valuation': 106, 'Industry Dashboard': 30, 'Sources': 13, 'Audit Checks': 13,
    'Scenario Manager': 68, 'Sensitivity Analysis': 95,
}
wb = openpyxl.load_workbook('incoming/IFM_user.xlsx')
for ws in wb.worksheets:
    lim = LEGACY.get(ws.title, ws.max_row + 1)
    fmts = {}
    n_empty = 0
    for r in range(1, min(lim, ws.max_row + 1)):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            if cell.value is not None:
                continue
            # only count cells that are inside a styled (bordered) region -> part of a designed table
            b = cell.border
            if not (b.left.style or b.right.style or b.top.style or b.bottom.style):
                continue
            n_empty += 1
            fmts[cell.number_format] = fmts.get(cell.number_format, 0) + 1
    if n_empty:
        top = sorted(fmts.items(), key=lambda kv: -kv[1])[:6]
        print(f"{ws.title:30s} bordered-empty={n_empty:5d}  " + "  ".join(f"{k!r}:{v}" for k, v in top))
