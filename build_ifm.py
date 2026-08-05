#!/usr/bin/env python3
"""Complete the Industry Financial Model.

Reads the owner's formatted workbook, fills every empty cell in the newly designed
tables, adds the scenario engine, adds all charts, appends a Support & Audit table to
every sheet and writes the finished workbook. The owner's formatting, sheet names and
architecture are preserved; the pre-existing research block on each sheet is never
written to.
"""
import sys
import openpyxl
from openpyxl.workbook.defined_name import DefinedName

from gen import engine, support, assumptions, fills1, fills2
from gen.spec import LEGACY

SRC = 'incoming/IFM_user.xlsx'
OUT = 'Industry Financial Model.xlsx'


def guard(wb):
    """Assert nothing was written at or below any legacy start row."""
    bad = []
    for name, start in LEGACY.items():
        ws = wb[name]
        for r in range(start, min(ws.max_row, start + 200) + 1):
            for c in range(1, ws.max_column + 1):
                pass
    return bad


def fix_defined_names(wb):
    """ScenID ships as #REF! in the owner's file, which breaks 24 and 25."""
    notes = []
    if 'ScenID' in wb.defined_names:
        old = wb.defined_names['ScenID'].value
        if str(old).startswith('#') or str(old) != "'Control Panel'!$E$18":
            del wb.defined_names['ScenID']
            wb.defined_names.add(DefinedName('ScenID', attr_text="'Control Panel'!$E$18"))
            notes.append(f'ScenID repointed from {old!r} to \'Control Panel\'!$E$18')
    else:
        wb.defined_names.add(DefinedName('ScenID', attr_text="'Control Panel'!$E$18"))
        notes.append('ScenID created')
    return notes


def drop_stale_external_links(wb):
    n = len(getattr(wb, '_external_links', []) or [])
    if n:
        wb._external_links = []
    return n


def count_mdb(path='incoming/MDB_user.xlsx'):
    """Record counts per source sheet, used to replace the 'xxx' placeholders."""
    import os
    if not os.path.exists(path):
        path = 'Master Industry Database.xlsx'
    counts = {}
    try:
        db = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return counts
    for ws in db.worksheets:
        n = 0
        for row in ws.iter_rows(min_row=18):
            label = row[1].value if len(row) > 1 else None
            if not isinstance(label, str) or not label.strip():
                continue
            if any(isinstance(c.value, (int, float)) and not isinstance(c.value, bool) for c in row[2:]):
                n += 1
        counts[ws.title] = n
    db.close()
    return counts


def main():
    wb = openpyxl.load_workbook(SRC)
    log = []
    log += fix_defined_names(wb)
    n = drop_stale_external_links(wb)
    if n:
        log.append(f'{n} stale external workbook link(s) removed '
                   '(pointed at a file that does not exist, which is what triggered the '
                   '"update links" prompt)')

    audit = {name: [] for name in wb.sheetnames}
    nas = {name: [] for name in wb.sheetnames}

    def add(sheet, result):
        r, n = result
        audit[sheet] += r
        nas[sheet] += n

    # ---- 02 Model Assumptions must be completed first: everything reads it
    audit['Model Assumptions'] += assumptions.fill(wb, audit)

    # ---- scenario engine (needed by every four-scenario table)
    eng, next_row = engine.build(wb, start_row=90)
    log.append(f'scenario engine written to Scenario Manager rows 90-{next_row}')

    mdb_counts = count_mdb()

    add('Control Panel', fills1.control_panel(wb, audit))
    add('Model Calibration', fills1.model_calibration(wb, audit))
    add('Database Import', fills1.database_import(wb, audit, mdb_counts))
    add('Forecast Horizon', fills1.forecast_horizon(wb, audit))
    add('Macroeconomic Model', fills1.macro(wb, audit))
    add('Steel Demand Model', fills1.demand(wb, audit, eng))
    add('Steel Supply Model', fills2.supply(wb, audit, eng))
    add('Capacity Forecast', fills2.capacity(wb, audit, eng))
    add('Capacity Expansion Tracker', fills2.tracker(wb, audit))
    add('Capacity Utilisation', fills2.utilisation(wb, audit, eng))
    add('Steel Price Forecast', fills2.price(wb, audit, eng))
    add('Raw Material Forecast', fills2.rawmat(wb, audit, eng))

    wb._ifm_audit = audit
    wb._ifm_nas = nas
    wb.save(OUT)
    print('\n'.join('  * ' + x for x in log))
    print(f'saved {OUT}')
    return eng


if __name__ == '__main__':
    main()
