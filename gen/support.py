"""Register entry producers.

Builds the dictionaries that gen/register.py turns into the reference register on
25 Audit Checks. Entries come from three places:

 1. the pre-existing research block (its own Line item / Unit / Source / Conf. /
    Notes columns are the authoritative provenance and are re-presented here),
 2. hand-written entries for outputs that live only in the new upper tables,
 3. one entry for every missing-data group, carrying the reason it is unavailable.

Every entry carries an `anchor` (the row or rows on its own sheet that it describes) so
the register can stamp its code in the Ref column beside those rows. The Value column is
a LIVE formula, re-qualified onto the origin sheet by register.qualify(), so the register
re-states itself when the scenario changes.

This module no longer writes anything into a worksheet: the per-sheet Support & Audit
tables it used to append have been replaced by the single central register.
"""
import re

from .spec import CODE, LEGACY

FREQ_BY_SOURCE = [
    (r'RBI|MPC|repo|G-sec|CRISIL|investing', 'On each RBI MPC statement (bi-monthly)', 'Jun-2026'),
    (r'World Bank|Pink', 'Monthly, on each Pink Sheet release', 'Jun-2026'),
    (r'worldsteel|OECD', 'Twice yearly (April and October SRO)', 'Apr-2026'),
    (r'Ministry|JPC|Joint Plant', 'Monthly, on each Ministry of Steel report', 'Apr-2026'),
    (r'Master DB|Database|Derived|Derived-DB|00 Database', 'Quarterly, on each results season', 'FY2026A'),
    (r'policy|P\d\d|Statutory|Convention', 'On any statutory or policy change', 'In force'),
    (r'Indicative|Not sourced|NOT SOURCED|Composite', 'Quarterly review; replace before transaction use', 'n/a'),
    (r'Computed|^\d\d ', 'Recalculates live with the model', 'Live'),
]


def _freq(src):
    s = str(src or '')
    for pat, freq, last in FREQ_BY_SOURCE:
        if re.search(pat, s, re.I):
            return freq, last
    return 'Quarterly', 'FY2026A'


def _method(formula, const_val):
    """Classify a cell into one of the model's field types."""
    if formula is None:
        return 'Sourced actual (hard value)' if const_val is not None else 'Not populated'
    t = formula if isinstance(formula, str) else getattr(formula, 'text', '')
    if not t.startswith('='):
        return 'Sourced actual (hard value)'
    if 'XLOOKUP' in t or 'INDEX(' in t and 'ScenID' in t:
        return 'Scenario-driven lookup'
    if 'SUMIFS' in t:
        return 'Database-driven aggregation (SUMIFS)'
    if "'" in t and '!' in t:
        return 'Cross-sheet reference (linked, not re-entered)'
    if re.search(r'\bMIN\(|\bMAX\(|\bIF\(', t):
        return 'Conditional / constrained calculation'
    return 'Computed in-cell'


def _sheets_in(formula):
    t = formula if isinstance(formula, str) else getattr(formula, 'text', '') or ''
    names = sorted(set(re.findall(r"'([^']+)'!", t)))
    return ', '.join(names) if names else 'This sheet'


def _split_note(note):
    """Split the legacy Notes column into assumption / reasoning / cross-check."""
    n = (note or '').strip()
    if not n:
        return '', '', ''
    parts = re.split(r'(?<=[.!])\s+', n)
    assumption = parts[0] if parts else ''
    reasoning = ' '.join(parts[1:3]) if len(parts) > 1 else ''
    cross = ' '.join(parts[3:]) if len(parts) > 3 else ''
    return assumption, reasoning, cross


def _formula_text(cell):
    v = cell.value
    if v is None:
        return ''
    if isinstance(v, str):
        return v if v.startswith('=') else f'(constant) {v}'[:250]
    if hasattr(v, 'text'):
        return '{=' + str(v.text).lstrip('=') + '}'
    return f'(constant) {v}'


# the research block on these sheets does not use the Line item / Unit / Source grammar
SPECIAL = {'Audit Checks', 'Sources'}
# where the research block stops, if something was appended below it
LEGACY_STOP = {'Scenario Manager': 87}


def harvest_special(ws):
    """25 Audit Checks and 26 Sources carry their own column grammar."""
    rows = []
    if ws.title == 'Audit Checks':
        for r in range(14, 45):
            cid = ws.cell(r, 1).value
            val = ws.cell(r, 2).value
            if not isinstance(cid, str) and not isinstance(val, str):
                continue
            rows.append(dict(
                item=f'{cid or "OVERALL"} - {val or "overall result"}',
                value=f'=IF(C{r}="PASS",1,IF(C{r}="WARN",0.5,0))', unit='PASS / WARN / FAIL',
                method='Live conditional test evaluated on every recalculation',
                formula=_formula_text(ws.cell(r, 3)),
                primary='Computed from the model itself',
                secondary='Several checks are duplicated independently on 02b Model Calibration',
                assumption='None. A check is a test, not an assumption.',
                reasoning=str(ws.cell(r, 4).value or '')[:900],
                cross='Re-evaluates when the scenario is changed on 01 Control Panel',
                conf='High', linked=_sheets_in(ws.cell(r, 3).value),
                freq='Live', last='Live', anchor=r,
                comments=f'Research block row {r}. A WARN is a signal to read the number, not a defect.',
                numfmt='0.0'))
        return rows
    for r in range(14, ws.max_row + 1):
        ref = ws.cell(r, 1).value
        if not isinstance(ref, str) or not ref.strip():
            continue
        rows.append(dict(
            item=f'{ref} - {ws.cell(r, 3).value or ""}', value=None,
            unit=f'Level {ws.cell(r, 2).value}' if ws.cell(r, 2).value else '',
            method='Source record', formula='n/a',
            primary=f'{ws.cell(r, 3).value or ""} - {ws.cell(r, 4).value or ""}',
            secondary=str(ws.cell(r, 7).value or ''),
            assumption=str(ws.cell(r, 6).value or ''),
            reasoning=str(ws.cell(r, 8).value or '')[:900],
            cross=str(ws.cell(r, 7).value or ''), conf='Level ' + str(ws.cell(r, 2).value or ''),
            linked='02 Model Assumptions', freq='On each publication cycle', anchor=r,
            last=str(ws.cell(r, 5).value or ''), comments=f'Source register row {r}.'))
    return rows


def harvest_legacy(ws):
    """Build support rows from the pre-existing research block."""
    start = LEGACY.get(ws.title)
    rows = []
    if not start:
        return rows
    if ws.title in SPECIAL:
        return harvest_special(ws)
    stop = LEGACY_STOP.get(ws.title, ws.max_row)
    for r in range(start, stop + 1):
        item = ws.cell(r, 1).value
        if not isinstance(item, str) or not item.strip():
            continue
        item = item.strip()
        unit = ws.cell(r, 2).value
        # header rows and section banners: no unit, no numbers -> skip banners but
        # keep them out of the table entirely
        if item in ('Line item',) or (unit in (None, '') and not any(
                ws.cell(r, c).value is not None for c in range(3, 11))):
            continue
        src = ws.cell(r, 12).value
        conf = ws.cell(r, 13).value
        note = ws.cell(r, 14).value
        # pick the cell that best represents the row
        fcell = None
        for c in (4, 3, 8, 5):
            cc = ws.cell(r, c)
            if cc.value is not None:
                fcell = cc
                break
        formula = fcell.value if fcell is not None else None
        has_a = ws.cell(r, 3).value is not None
        has_t = ws.cell(r, 10).value is not None
        if has_t and has_a:
            value = f'=IF(ISNUMBER(J{r}),J{r},C{r})'
        elif has_t:
            value = f'=J{r}'
        elif has_a:
            value = f'=C{r}'
        else:
            value = f'=IFERROR(D{r},"")'
        a, rea, cr = _split_note(note)
        freq, last = _freq(src)
        rows.append(dict(
            item=item, value=value, unit=unit or '', method=_method(formula, ws.cell(r, 3).value),
            formula=_formula_text(fcell) if fcell is not None else '', primary=src or 'See 26 Sources',
            secondary='26 Sources register; Master Industry Database source log (47 cited sources)',
            assumption=a, reasoning=rea, cross=cr or 'Reconciled by 25 Audit Checks',
            conf=conf or 'Medium', linked=_sheets_in(formula), freq=freq, last=last, anchor=r,
            comments=f'Research block row {r}. Value column is live and re-states on a scenario change.',
            numfmt=None))
    return rows


def entries(ws, extra_rows=None, na_rows=None):
    """All register entries for `ws`, in presentation order."""
    return list(extra_rows or []) + harvest_legacy(ws) + list(na_rows or [])


INTRO = {
    'Cover': 'This workbook is an integrated industry financial model, not a set of independent '
             'forecasts. Real GDP growth and a demand elasticity drive volume; volume against capacity '
             'drives utilisation; utilisation feeds back into realisation with a one-year lag; '
             'realisation less cash cost drives margin, cash flow and valuation. The table below records '
             'the coverage universe and the model metadata on this page.',
    'Index': 'Every sheet in the workbook, its role, and where its numbers come from. Sheet names in '
             'column B are live hyperlinks.',
    'Database Import': 'What is imported from Master Industry Database.xlsx, how many records each '
                       'dataset carries, and the external reference formula that should replace each '
                       'RED cell when the two workbooks are reconnected.',
    'Forecast Horizon': 'Why the forecast horizon is seven years. The scoring is a framework; the '
                        'evidence behind each score is stated on the sheet.',
    'Control Panel': 'The six workbook status tests and the fourteen navigation links. Every status '
                     'cell is a live formula that can fail - none is a hard-coded tick.',
    'Model Assumptions': 'THE SINGLE SOURCE OF TRUTH FOR THE WHOLE WORKBOOK. Twenty-six drivers, each '
                         'with four scenario paths and a documented source. Nothing downstream re-keys '
                         'any of these values; everything links here. The table below documents the '
                         'FY2026A anchor of every driver.',
    'Model Calibration': 'The FY2026A actuals every forecast is anchored on, and five arithmetic '
                         'reconciliations that prove the anchors are internally consistent.',
    'Audit Checks': 'Thirty live validations plus the standing limitations. The table below explains '
                    'what each check proves and what to do if it fails.',
    'Sources': 'Every source used for a forward-looking assumption, with publisher, publication, date '
               'and what it is used for.',
}


def na_row(label, addrs, reason, unit='', linked='This sheet'):
    """Register entry describing a MISSING-DATA (red) group.

    Reserved for cells that can be neither sourced nor defensibly modelled. `addrs` is
    kept as structured data on the entry so the register can stamp the code against every
    affected row and so tools/audit.py can reconcile the addresses against the actual
    red cells instead of trusting prose.
    """
    return dict(item=label, value=None, unit=unit, addrs=addrs,
                method='NOT POPULATED - red fill (missing data)', formula='(blank by design)',
                primary='No reliable source located and no defensible basis for an estimate',
                secondary='n/a',
                assumption='None asserted. The cell is deliberately blank.',
                reasoning=reason, cross='Excluded from every headline output and from the audit checks',
                conf='n/a - unavailable', linked=linked,
                freq='Populate when a primary source becomes available',
                last='Not available', comments=f'Red (missing) cells: {addrs}')


def est_row(label, addrs, basis, reason, unit='', linked='This sheet', value=None,
            primary='Modelled estimate - see Assumption and Reasoning', secondary='n/a',
            conf='Medium - modelled, not sourced', formula='', method='', numfmt=None,
            cross='', freq='Review annually, or when a published series becomes available',
            last='FY2026A', comments=''):
    """Register entry for a block that is MODELLED rather than sourced.

    Used where no published series exists but the value follows from a stated driver. The
    cell is populated so the model is complete; the assumption, its basis and its
    confidence are recorded here so the number can be challenged.
    """
    return dict(item=label, value=value, unit=unit, addrs=addrs,
                method=method or 'Modelled from a documented driver, not sourced',
                formula=formula or '(see Calculation Method)',
                primary=primary, secondary=secondary,
                assumption=basis, reasoning=reason,
                cross=cross or 'Recomputes with the driver it is built on, so it cannot drift',
                conf=conf, linked=linked, freq=freq, last=last,
                comments=comments or 'Populated by modelling rather than left blank; the basis is stated '
                                     'in the Assumption column so it can be replaced with a sourced '
                                     'series when one exists.',
                numfmt=numfmt)
