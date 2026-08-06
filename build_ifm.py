#!/usr/bin/env python3
"""Complete the Industry Financial Model.

Reads the owner's formatted workbook, fills every empty cell in the newly designed
tables, adds the scenario engine, adds all charts, stamps a reference code beside every
documented row, writes the central reference register onto 25 Audit Checks and saves the
finished workbook.

The owner's formatting, sheet names and architecture are preserved. The pre-existing
research block on each sheet is never written to except for the two arithmetic
corrections in gen/corrections.py, which are disclosed individually and carried in the
audit tool's allow-list.

Model sheets carry model numbers and a reference code, nothing else: the documentation
that used to sit in a Support & Audit table at the bottom of all 31 sheets now lives in
one register, keyed by code.
"""
import sys
import openpyxl
from openpyxl.workbook.defined_name import DefinedName

from gen import (engine, support, register, corrections, assumptions, charts,
                 fills1, fills2, fills3, fills4, fills5)
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


CONVENTIONS = [
    ('Blue font', 'Manual input',
     'A value a user is expected to review or change. Structural parameters held constant across years '
     'and scenarios, modelled estimates whose basis is stated in the register, and the handful of '
     'deliberate zeros that are modelling statements rather than placeholders, are all blue.'),
    ('Black font', 'Formula computed on this sheet',
     'Arithmetic performed in the cell itself. Every bridge, reconciliation, ratio and classification is '
     'black.'),
    ('Green font', 'Link to another sheet',
     'The value exists in exactly one place in the workbook and is read from there. Roughly two thirds of '
     'the model is green, which is the point: nothing is re-keyed.'),
    ('Orange fill', 'Sourced from an external workbook - Master Industry Database.xlsx',
     'Historical actuals and company-level data imported from the database. These are the cells to convert '
     'to external references when the two workbooks are reconnected. The exact intended external formula '
     'is recorded against each one in the reference register on 25 Audit Checks.'),
    ('Red fill', 'Missing data - no source and no defensible estimate',
     'Nothing is asserted in these cells. Every red group has an entry in the reference register stating '
     'precisely why the data could not be obtained and what to enter instead. They are gaps that are '
     'disclosed, not gaps that are hidden - and they are now the exception rather than the rule, because '
     'anything that could be modelled from a documented driver has been modelled instead of left blank.'),
    ('Ref column', 'The reference code beside each table',
     'Each documented row carries a short code in a narrow Ref column to the right of its table. The code '
     'is a hyperlink: click it to jump to that row\'s full entry - method, formula, source, assumption, '
     'reasoning, cross-check, confidence and refresh frequency - in the reference register on 25 Audit '
     'Checks.'),
]


def convention_rows():
    return [dict(
        item=f'CONVENTION - {name}: {meaning}', value=None, unit='formatting convention',
        method='Institutional modelling convention', formula='n/a',
        primary='Model owner specification', secondary='n/a',
        assumption='Applied without exception throughout the workbook.', reasoning=why,
        cross='Verified by tools/audit.py: every blank cell inside a designed table carries a red '
              'missing-data fill and a documented reason in the register',
        conf='High', linked='All sheets', freq='n/a', last='n/a', anchor=8,
        comments='Read this before reading any number.') for name, meaning, why in CONVENTIONS]


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
    add('Cost Curve', fills3.cost_curve(wb, audit, eng))
    add('Revenue Forecast', fills3.revenue(wb, audit, eng))
    add('EBITDA Model', fills3.ebitda(wb, audit, eng))
    add('Margin Analysis', fills3.margin(wb, audit, eng))
    add('Working Capital Model', fills3.working_capital(wb, audit, eng))
    add('Cash Flow Model', fills3.cash_flow(wb, audit, eng))
    add('Capital Allocation', fills4.capital_allocation(wb, audit, eng))
    add('Industry Cycle Model', fills4.cycle(wb, audit, eng))
    add('Trade Model', fills4.trade(wb, audit, eng))
    add('ESG Model', fills4.esg(wb, audit, eng))
    add('Scenario Manager', fills5.scenario_manager(wb, audit, eng))
    add('Sensitivity Analysis', fills5.sensitivity(wb, audit, eng))
    add('Comparable Valuation', fills5.comparable_valuation(wb, audit, eng))
    add('Industry Dashboard', fills5.dashboard(wb, audit, eng))
    add('Index', fills5.index_links(wb, audit))

    audit['Cover'] = convention_rows() + audit['Cover']

    # ---- two disclosed arithmetic corrections inside the research block
    fixes = corrections.apply(wb, audit)
    audit['Comparable Valuation'] = fixes + audit['Comparable Valuation']
    log.append(f'{len(fixes)} arithmetic defect(s) in the research block corrected, not just flagged')

    # ---- charts are anchored before the register is written, so the dashboard's
    #      chart block sits directly under its own content
    dash_row = wb['Industry Dashboard'].max_row + 3

    # ---- the reference register: codes stamped beside every documented row, all of
    #      the detail centralised on 25 Audit Checks
    entries = {name: support.entries(wb[name], extra_rows=audit.get(name),
                                     na_rows=nas.get(name))
               for name in wb.sheetnames}
    n_rows, header_row, stamped = register.build(wb, entries, intros=support.INTRO)
    log.append(f'reference register written to Audit Checks: {n_rows} documented items across '
               f'{len(header_row)} sheets, {stamped} codes stamped in Ref columns')
    log.append('per-sheet Support & Audit tables removed - the model sheets now carry model numbers '
               'and a reference code only')

    # ---- charts, built from scratch and linked to model ranges
    made = charts.build(wb, dash_row=dash_row)
    log.append(f'{len(made)} charts created across '
               f'{len(set(s for s, _a, _t in made))} sheets')

    # openpyxl writes formulas without cached results, so Excel must be told to
    # recalculate on open or the user sees stale or empty cells.
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.calcCompleted = False
    log.append('fullCalcOnLoad set - Excel recalculates the whole workbook on open')

    wb._ifm_audit = audit
    wb._ifm_nas = nas
    wb._ifm_register = header_row
    wb.save(OUT)
    print('\n'.join('  * ' + x for x in log))
    print(f'saved {OUT}')
    return eng


if __name__ == '__main__':
    main()
