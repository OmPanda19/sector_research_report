"""Charts for the Industry Financial Model.

Every chart is built from scratch and every series is a Reference into a model range,
so nothing is hard-coded and every chart re-draws when the scenario is changed on
01 Control Panel. Formatting follows one institutional palette - no default Excel
colours, no default Excel style numbers.

Layout convention in this workbook:
  * new upper tables put YEARS ACROSS COLUMNS  -> series are rows      (rows=True)
  * comparison tables put ITEMS DOWN ROWS      -> series are columns   (rows=False)
Getting that wrong silently transposes a chart, so every entry below states it.
"""
from openpyxl.chart import BarChart, LineChart, PieChart, DoughnutChart, RadarChart, Reference
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import Marker
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (Paragraph, ParagraphProperties, CharacterProperties,
                                   Font as DFont)

NAVY = '1F3864'
STEEL = '2E5C8A'
TEAL = '2E8B8B'
GOLD = 'C9A227'
BURGUNDY = '8B2E3C'
GREY = '9C9C9C'
SKY = '7FA8D0'
SAGE = '8FA98F'
CHARCOAL = '4A4A4A'
PALETTE = [NAVY, STEEL, GOLD, TEAL, BURGUNDY, SKY, GREY, SAGE]
GRID = 'DDDDDD'

PCT, INT, ONE, X, TWO = '0.0%', '#,##0', '#,##0.0', '0.0"x"', '0.00'


def _cp(sz=900, b=False, color=CHARCOAL):
    return CharacterProperties(latin=DFont(typeface='Calibri'), sz=sz, b=b, solidFill=color)


def _rt(sz=900, b=False, color=CHARCOAL):
    pr = ParagraphProperties(defRPr=_cp(sz, b, color))
    return RichText(p=[Paragraph(pPr=pr, endParaRPr=_cp(sz, b, color))])


def _title(ch, text):
    ch.title = text
    try:
        ch.title.tx.rich.p[0].pPr = ParagraphProperties(defRPr=_cp(1050, True, NAVY))
    except (AttributeError, IndexError, TypeError):
        pass


def _axis_title(ax, text):
    if not text:
        return
    ax.title = text
    try:
        ax.title.tx.rich.p[0].pPr = ParagraphProperties(defRPr=_cp(850, True, GREY))
    except (AttributeError, IndexError, TypeError):
        pass


def _frame(ch, title, xt, yt, yfmt, legend, w, h, grid=True):
    _title(ch, title)
    ch.style = None
    ch.width, ch.height = w, h
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.x_axis.majorTickMark = 'out'
    ch.y_axis.majorTickMark = 'none'
    ch.x_axis.txPr = _rt(800)
    ch.y_axis.txPr = _rt(800)
    ch.y_axis.numFmt = yfmt
    ch.y_axis.majorGridlines = ChartLines(
        spPr=GraphicalProperties(ln=LineProperties(solidFill=GRID, w=3175))) if grid else None
    ch.x_axis.spPr = GraphicalProperties(ln=LineProperties(solidFill=GREY, w=6350))
    ch.y_axis.spPr = GraphicalProperties(ln=LineProperties(noFill=True))
    _axis_title(ch.x_axis, xt)
    _axis_title(ch.y_axis, yt)
    if legend and ch.legend is not None:
        ch.legend.position = legend
        ch.legend.overlay = False
        ch.legend.txPr = _rt(800)
    else:
        ch.legend = None
    return ch


def _colour_series(ch, colors, line=False, markers=True):
    for i, s in enumerate(ch.series):
        col = colors[i % len(colors)]
        if line:
            s.graphicalProperties = GraphicalProperties(
                ln=LineProperties(solidFill=col, w=22225))
            s.smooth = False
            s.marker = Marker(symbol='circle', size=5,
                              spPr=GraphicalProperties(
                                  solidFill=col,
                                  ln=LineProperties(solidFill='FFFFFF', w=6350))) \
                if markers else Marker(symbol='none')
        else:
            s.graphicalProperties = GraphicalProperties(solidFill=col,
                                                        ln=LineProperties(noFill=True))


def _labels(ch, fmt):
    ch.dataLabels = DataLabelList()
    ch.dataLabels.showVal = True
    ch.dataLabels.numFmt = fmt
    ch.dataLabels.txPr = _rt(750)


def make(ws, spec):
    """Build one chart from a spec dict."""
    kind = spec['kind']
    rows = spec.get('rows', True)
    tfd = spec.get('titles', True)
    colors = spec.get('colors', PALETTE)
    yfmt = spec.get('yfmt', INT)
    w, h = spec.get('w', 18.0), spec.get('h', 8.8)
    legend = spec.get('legend', 'b')
    cats = Reference(ws, *spec['cats']) if spec.get('cats') else None

    if kind == 'combo':
        b = BarChart()
        b.type = 'col'
        b.gapWidth = spec.get('gap', 70)
        b.add_data(Reference(ws, *spec['bar']), titles_from_data=tfd, from_rows=rows)
        if cats is not None:
            b.set_categories(cats)
        _colour_series(b, spec.get('bar_colors', [STEEL]))
        _frame(b, spec['title'], spec.get('xt'), spec.get('yt'), yfmt, legend, w, h)
        ln = LineChart()
        ln.add_data(Reference(ws, *spec['line']), titles_from_data=tfd, from_rows=rows)
        _colour_series(ln, spec.get('line_colors', [GOLD]), line=True)
        ln.y_axis.axId = 200
        ln.y_axis.numFmt = spec.get('y2fmt', PCT)
        ln.y_axis.delete = False
        ln.y_axis.majorGridlines = None
        ln.y_axis.txPr = _rt(800)
        ln.y_axis.crosses = 'max'
        _axis_title(ln.y_axis, spec.get('y2t'))
        b += ln
        return b

    if kind in ('bar', 'hbar', 'stack'):
        ch = BarChart()
        ch.type = 'bar' if kind == 'hbar' else 'col'
        if kind == 'stack':
            ch.grouping = 'stacked'
            ch.overlap = 100
        ch.gapWidth = spec.get('gap', 60)
        ch.add_data(Reference(ws, *spec['data']), titles_from_data=tfd, from_rows=rows)
        if cats is not None:
            ch.set_categories(cats)
        _colour_series(ch, colors)
        if spec.get('labels'):
            _labels(ch, spec.get('lblfmt', yfmt))
        return _frame(ch, spec['title'], spec.get('xt'), spec.get('yt'), yfmt, legend, w, h)

    if kind == 'line':
        ch = LineChart()
        ch.add_data(Reference(ws, *spec['data']), titles_from_data=tfd, from_rows=rows)
        if cats is not None:
            ch.set_categories(cats)
        _colour_series(ch, colors, line=True, markers=spec.get('markers', True))
        return _frame(ch, spec['title'], spec.get('xt'), spec.get('yt'), yfmt, legend, w, h)

    if kind in ('pie', 'doughnut'):
        ch = DoughnutChart(holeSize=52) if kind == 'doughnut' else PieChart()
        ch.add_data(Reference(ws, *spec['data']), titles_from_data=tfd, from_rows=rows)
        if cats is not None:
            ch.set_categories(cats)
        n = spec.get('points', 8)
        ch.series[0].data_points = [
            DataPoint(idx=i, spPr=GraphicalProperties(
                solidFill=colors[i % len(colors)],
                ln=LineProperties(solidFill='FFFFFF', w=12700))) for i in range(n)]
        ch.dataLabels = DataLabelList()
        ch.dataLabels.showPercent = True
        ch.dataLabels.txPr = _rt(750, color='FFFFFF')
        _title(ch, spec['title'])
        ch.width, ch.height = w, h
        ch.legend.position = 'r'
        ch.legend.txPr = _rt(750)
        return ch

    if kind == 'radar':
        ch = RadarChart()
        ch.type = 'marker'
        ch.add_data(Reference(ws, *spec['data']), titles_from_data=tfd, from_rows=rows)
        if cats is not None:
            ch.set_categories(cats)
        _colour_series(ch, colors, line=True)
        _title(ch, spec['title'])
        ch.width, ch.height = w, h
        ch.x_axis.txPr = _rt(750)
        ch.y_axis.txPr = _rt(750)
        ch.y_axis.delete = False
        ch.y_axis.numFmt = yfmt
        ch.legend.position = 'b'
        ch.legend.txPr = _rt(800)
        return ch

    raise ValueError(kind)


# =====================================================================================
# (sheet, anchor, spec).  Reference args are (min_col, min_row, max_col, max_row).
# =====================================================================================
Y8 = (2, 8, 9, 8)          # new-table year header B..I on most sheets

SPECS = [
    # ---------------------------------------------------------- 04 Steel Demand Model
    ('Steel Demand Model', 'K7', dict(
        kind='combo', title='Apparent Finished Steel Consumption and Growth',
        bar=(1, 16, 9, 16), line=(1, 17, 9, 17), cats=(2, 15, 9, 15), rows=True,
        bar_colors=[STEEL], line_colors=[GOLD], yfmt=INT, y2fmt=PCT,
        yt='Consumption (Mt)', y2t='Growth (%)')),
    ('Steel Demand Model', 'K22', dict(
        kind='line', title='Per Capita Finished Steel Consumption',
        data=(1, 18, 9, 18), cats=(2, 15, 9, 15), rows=True, colors=[TEAL],
        yfmt=INT, yt='kg per capita', legend=None)),
    ('Steel Demand Model', 'K37', dict(
        kind='bar', title='Per Capita Consumption: India against the World (kg)',
        data=(2, 72, 2, 78), cats=(1, 73, 1, 78), rows=False,
        colors=[NAVY, BURGUNDY, GREY, GREY, GREY, STEEL], yfmt=INT, legend=None,
        labels=True)),
    ('Steel Demand Model', 'K52', dict(
        kind='line', title='Apparent Consumption by Scenario (Mt)',
        data=(2, 93, 5, 100), cats=(1, 94, 1, 100), rows=False,
        colors=[NAVY, TEAL, GOLD, BURGUNDY], yfmt=INT, yt='Mt')),

    # ---------------------------------------------------------- 05 Steel Supply Model
    ('Steel Supply Model', 'K7', dict(
        kind='line', title='Crude Production, Finished Production and Installed Capacity',
        data=(1, 9, 9, 11), cats=Y8, rows=True, colors=[STEEL, TEAL, NAVY],
        yfmt=INT, yt='Mt / Mtpa')),
    ('Steel Supply Model', 'K22', dict(
        kind='combo', title='Domestic Availability and Capacity Utilisation',
        bar=(1, 13, 9, 13), line=(1, 12, 9, 12), cats=Y8, rows=True,
        bar_colors=[SKY], line_colors=[BURGUNDY], yfmt=INT, y2fmt=PCT,
        yt='Availability (Mt)', y2t='Utilisation (%)')),
    ('Steel Supply Model', 'K37', dict(
        kind='bar', title='Historical Production, FY2021-FY2026 (Mt)',
        data=(2, 20, 3, 26), cats=(1, 21, 1, 26), rows=False,
        colors=[NAVY, STEEL], yfmt=INT)),

    # ---------------------------------------------------------- 06 Capacity Forecast
    ('Capacity Forecast', 'K7', dict(
        kind='combo', title='Installed Capacity and Utilisation',
        bar=(1, 9, 9, 9), line=(1, 12, 9, 12), cats=Y8, rows=True,
        bar_colors=[NAVY], line_colors=[GOLD], yfmt=INT, y2fmt=PCT,
        yt='Capacity (Mtpa)', y2t='Utilisation (%)')),
    ('Capacity Forecast', 'K22', dict(
        kind='stack', title='Capacity Additions by Type (Mtpa)',
        data=(1, 52, 8, 54), cats=(2, 50, 8, 50), rows=True,
        colors=[STEEL, TEAL, SKY], yfmt=ONE)),
    ('Capacity Forecast', 'K37', dict(
        kind='line', title='Capacity against Demand-Driven Requirement (Mtpa)',
        data=(2, 61, 3, 68), cats=(1, 62, 1, 68), rows=False,
        colors=[NAVY, BURGUNDY], yfmt=INT, yt='Mtpa')),

    # ---------------------------------------------------------- 08 Capacity Utilisation
    ('Capacity Utilisation', 'K7', dict(
        kind='line', title='Industry Capacity Utilisation',
        data=(1, 11, 9, 11), cats=Y8, rows=True, colors=[NAVY], yfmt=PCT,
        yt='Utilisation (%)', legend=None)),
    ('Capacity Utilisation', 'K22', dict(
        kind='bar', title='Installed Capacity by Producer, FY2026A (Mtpa)',
        data=(7, 18, 7, 26), cats=(6, 19, 6, 26), rows=False, colors=[NAVY],
        yfmt=ONE, legend=None, labels=True)),
    ('Capacity Utilisation', 'K37', dict(
        kind='line', title='Capacity Utilisation by Scenario',
        data=(2, 51, 5, 58), cats=(1, 52, 1, 58), rows=False,
        colors=[NAVY, TEAL, GOLD, BURGUNDY], yfmt=PCT, yt='Utilisation (%)')),

    # ---------------------------------------------------------- 09 Steel Price Forecast
    ('Steel Price Forecast', 'K7', dict(
        kind='line', title='Blended Industry Realisation (Rs/t)',
        data=(1, 12, 9, 12), cats=Y8, rows=True, colors=[NAVY], yfmt=INT,
        yt='Rs per tonne', legend=None)),
    ('Steel Price Forecast', 'K22', dict(
        kind='line', title='Blended Realisation by Scenario (Rs/t)',
        data=(2, 67, 5, 74), cats=(1, 68, 1, 74), rows=False,
        colors=[NAVY, TEAL, GOLD, BURGUNDY], yfmt=INT, yt='Rs per tonne')),
    ('Steel Price Forecast', 'K37', dict(
        kind='bar', title='FY2026A Cost Floor and Price Build (Rs/t)',
        data=(2, 45, 2, 54), cats=(1, 46, 1, 54), rows=False,
        colors=[BURGUNDY, CHARCOAL, SAGE, GREY, GREY, STEEL, TEAL, GOLD, NAVY],
        yfmt=INT, legend=None)),

    # ---------------------------------------------------------- 10 Raw Material Forecast
    ('Raw Material Forecast', 'Q7', dict(
        kind='combo', title='Coking Coal and Iron Ore',
        bar=(1, 10, 9, 10), line=(1, 9, 9, 9), cats=Y8, rows=True,
        bar_colors=[CHARCOAL], line_colors=[BURGUNDY], yfmt=INT, y2fmt=INT,
        yt='Coking coal (US$/t)', y2t='Iron ore (US$/dmt)')),
    ('Raw Material Forecast', 'Q22', dict(
        kind='line', title='Iron Ore 62% Fe CFR China, FY2021-FY2026 (US$/dmt)',
        data=(1, 21, 7, 21), cats=(2, 20, 7, 20), rows=True, colors=[BURGUNDY],
        yfmt=INT, yt='US$/dmt', legend=None)),
    ('Raw Material Forecast', 'Q37', dict(
        kind='bar', title='Input Costs by Scenario, FY2033E',
        data=(1, 66, 5, 67), cats=(2, 65, 5, 65), rows=True,
        colors=[BURGUNDY, CHARCOAL], yfmt=INT)),

    # ---------------------------------------------------------- 11 Cost Curve
    ('Cost Curve', 'K7', dict(
        kind='bar', title='Producer Cash Cost Curve, FY2026A (Rs/t)',
        data=(4, 18, 4, 24), cats=(2, 19, 2, 24), rows=False, colors=[TEAL],
        yfmt=INT, legend=None, labels=True)),
    ('Cost Curve', 'K22', dict(
        kind='doughnut', title='FY2026A Cash Cost Composition',
        data=(2, 37, 2, 47), cats=(1, 38, 1, 47), rows=False, points=10,
        colors=[BURGUNDY, CHARCOAL, GREY, GREY, SAGE, GREY, GREY, GREY, GREY, STEEL],
        w=13.5, h=9.0)),
    ('Cost Curve', 'K38', dict(
        kind='line', title='Cash Cost, Total Cost, Realisation and EBITDA per Tonne (Rs/t)',
        data=(1, 9, 9, 12), cats=Y8, rows=True, colors=[TEAL, SAGE, NAVY, GOLD],
        yfmt=INT, yt='Rs per tonne')),

    # ---------------------------------------------------------- 12 Revenue Forecast
    ('Revenue Forecast', 'K7', dict(
        kind='combo', title='Industry Revenue and Growth',
        bar=(1, 11, 9, 11), line=(1, 12, 9, 12), cats=Y8, rows=True,
        bar_colors=[NAVY], line_colors=[GOLD], yfmt=INT, y2fmt=PCT,
        yt='Revenue (Rs cr)', y2t='Growth (%)')),
    ('Revenue Forecast', 'K22', dict(
        kind='stack', title='Revenue Bridge: Prior Year, Volume Effect and Price Effect (Rs cr)',
        data=(1, 28, 8, 30), cats=(2, 27, 8, 27), rows=True,
        colors=[GREY, STEEL, GOLD], yfmt=INT)),
    ('Revenue Forecast', 'K38', dict(
        kind='stack', title='Domestic against Export Revenue (Rs cr)',
        data=(2, 51, 3, 59), cats=(1, 52, 1, 59), rows=False,
        colors=[NAVY, GOLD], yfmt=INT)),
    ('Revenue Forecast', 'K53', dict(
        kind='line', title='Industry Revenue by Scenario (Rs cr)',
        data=(2, 82, 5, 89), cats=(1, 83, 1, 89), rows=False,
        colors=[NAVY, TEAL, GOLD, BURGUNDY], yfmt=INT, yt='Rs crore')),

    # ---------------------------------------------------------- 13 EBITDA Model
    ('EBITDA Model', 'K7', dict(
        kind='combo', title='Industry EBITDA and Margin',
        bar=(1, 10, 9, 10), line=(1, 11, 9, 11), cats=Y8, rows=True,
        bar_colors=[NAVY], line_colors=[GOLD], yfmt=INT, y2fmt=PCT,
        yt='EBITDA (Rs cr)', y2t='Margin (%)')),
    ('EBITDA Model', 'K22', dict(
        kind='stack', title='EBITDA Bridge (Rs cr)',
        data=(1, 47, 8, 54), cats=(2, 46, 8, 46), rows=True,
        colors=[GREY, STEEL, GOLD, BURGUNDY, GREY, GREY, TEAL, SAGE], yfmt=INT, h=9.6)),
    ('EBITDA Model', 'K38', dict(
        kind='line', title='Margin Walk: Gross, EBITDA, Operating and Cash',
        data=(1, 81, 8, 84), cats=(2, 80, 8, 80), rows=True,
        colors=[SAGE, NAVY, STEEL, TEAL], yfmt=PCT, yt='Margin (%)')),
    ('EBITDA Model', 'K54', dict(
        kind='line', title='Industry EBITDA by Scenario (Rs cr)',
        data=(2, 88, 5, 95), cats=(1, 89, 1, 95), rows=False,
        colors=[NAVY, TEAL, GOLD, BURGUNDY], yfmt=INT, yt='Rs crore')),

    # ---------------------------------------------------------- 14 Margin Analysis
    ('Margin Analysis', 'K7', dict(
        kind='line', title='Margin Ladder: Gross, EBITDA, Operating and Cash',
        data=(1, 9, 9, 12), cats=Y8, rows=True, colors=[SAGE, NAVY, STEEL, TEAL],
        yfmt=PCT, yt='Margin (%)')),
    ('Margin Analysis', 'K22', dict(
        kind='stack', title='Margin Bridge (percentage points)',
        data=(1, 31, 8, 36), cats=(2, 30, 8, 30), rows=True,
        colors=[GREY, GOLD, BURGUNDY, GREY, GREY, TEAL], yfmt=PCT, h=9.6)),
    ('Margin Analysis', 'K38', dict(
        kind='bar', title='EBITDA Margin by Producer, FY2026A',
        data=(2, 55, 2, 60), cats=(1, 56, 1, 60), rows=False,
        colors=[NAVY, STEEL, GOLD, TEAL, GREY], yfmt=PCT, legend=None, labels=True)),
    ('Margin Analysis', 'K53', dict(
        kind='line', title='Industry EBITDA Margin by Scenario',
        data=(2, 75, 5, 82), cats=(1, 76, 1, 82), rows=False,
        colors=[NAVY, TEAL, GOLD, BURGUNDY], yfmt=PCT, yt='Margin (%)')),

    # ---------------------------------------------------------- 15 Working Capital Model
    ('Working Capital Model', 'K7', dict(
        kind='combo', title='Net Working Capital and Intensity',
        bar=(1, 12, 9, 12), line=(1, 13, 9, 13), cats=Y8, rows=True,
        bar_colors=[SKY], line_colors=[BURGUNDY], yfmt=INT, y2fmt=PCT,
        yt='NWC (Rs cr)', y2t='% of revenue')),
    ('Working Capital Model', 'K22', dict(
        kind='bar', title='Cash Absorbed by Working Capital (Rs cr)',
        data=(4, 46, 4, 53), cats=(1, 47, 1, 53), rows=False, colors=[BURGUNDY],
        yfmt=INT, legend=None, labels=True)),
    ('Working Capital Model', 'K37', dict(
        kind='line', title='Net Working Capital by Scenario (Rs cr)',
        data=(2, 57, 5, 64), cats=(1, 58, 1, 64), rows=False,
        colors=[NAVY, TEAL, GOLD, BURGUNDY], yfmt=INT, yt='Rs crore')),

    # ---------------------------------------------------------- 16 Cash Flow Model
    ('Cash Flow Model', 'K7', dict(
        kind='combo', title='Unlevered Free Cash Flow and Cash Conversion',
        bar=(1, 12, 9, 12), line=(1, 14, 9, 14), cats=Y8, rows=True,
        bar_colors=[TEAL], line_colors=[GOLD], yfmt=INT, y2fmt=PCT,
        yt='FCF (Rs cr)', y2t='Conversion (% of EBITDA)')),
    ('Cash Flow Model', 'K22', dict(
        kind='stack', title='Cash Flow Build-up (Rs cr)',
        data=(1, 20, 9, 25), cats=(2, 19, 9, 19), rows=True,
        colors=[NAVY, GREY, SKY, STEEL, TEAL, GOLD], yfmt=INT, h=9.6)),
    ('Cash Flow Model', 'K38', dict(
        kind='bar', title='EBITDA, Operating Cash Flow and Capex (Rs cr)',
        data=(1, 9, 9, 11), cats=Y8, rows=True, colors=[NAVY, STEEL, BURGUNDY],
        yfmt=INT)),
    ('Cash Flow Model', 'K54', dict(
        kind='line', title='Unlevered Free Cash Flow by Scenario (Rs cr)',
        data=(2, 78, 5, 85), cats=(1, 79, 1, 85), rows=False,
        colors=[NAVY, TEAL, GOLD, BURGUNDY], yfmt=INT, yt='Rs crore')),

    # ---------------------------------------------------------- 17 Capital Allocation
    ('Capital Allocation', 'K7', dict(
        kind='stack', title='Capex Allocation: Maintenance, Brownfield and Greenfield (Rs cr)',
        data=(1, 36, 8, 38), cats=(2, 35, 8, 35), rows=True,
        colors=[SKY, STEEL, NAVY], yfmt=INT)),
    ('Capital Allocation', 'K22', dict(
        kind='combo', title='Net Debt and Leverage',
        bar=(5, 46, 5, 53), line=(6, 46, 6, 53), cats=(1, 47, 1, 53), rows=False,
        bar_colors=[CHARCOAL], line_colors=[BURGUNDY], yfmt=INT, y2fmt=X,
        yt='Net debt (Rs cr)', y2t='Net debt / EBITDA')),
    ('Capital Allocation', 'K38', dict(
        kind='bar', title='Cumulative Growth Capex against Free Cash Flow by Scenario (Rs cr)',
        data=(2, 67, 3, 71), cats=(1, 68, 1, 71), rows=False,
        colors=[STEEL, TEAL], yfmt=INT)),

    # ---------------------------------------------------------- 18 Industry Cycle Model
    ('Industry Cycle Model', 'L7', dict(
        kind='line', title='Industry Cycle Score (0-100)',
        data=(3, 27, 10, 27), cats=(3, 18, 10, 18), rows=True, titles=False,
        colors=[NAVY], yfmt='0', yt='Score', legend=None)),
    ('Industry Cycle Model', 'L22', dict(
        kind='radar', title='Cycle Indicator Scores by Year',
        data=(3, 18, 10, 26), cats=(1, 19, 1, 26), rows=False, colors=PALETTE,
        yfmt='0', w=14.5, h=10.5)),
    ('Industry Cycle Model', 'L40', dict(
        kind='bar', title='FY2033E Cycle Score by Scenario',
        data=(2, 70, 5, 70), cats=(2, 69, 5, 69), rows=True, titles=False,
        colors=[NAVY, TEAL, GOLD, BURGUNDY], yfmt='0', legend=None, labels=True)),

    # ---------------------------------------------------------- 19 Trade Model
    ('Trade Model', 'K7', dict(
        kind='bar', title='Finished Steel Imports and Exports (Mt)',
        data=(1, 9, 9, 10), cats=Y8, rows=True, colors=[BURGUNDY, TEAL], yfmt=ONE)),
    ('Trade Model', 'K22', dict(
        kind='line', title='Import Penetration and Export Intensity',
        data=(1, 12, 9, 13), cats=Y8, rows=True, colors=[BURGUNDY, TEAL], yfmt=PCT,
        yt='% of consumption / production')),
    ('Trade Model', 'K37', dict(
        kind='bar', title='Historical Trade, FY2021-FY2026 (Mt)',
        data=(2, 18, 3, 24), cats=(1, 19, 1, 24), rows=False,
        colors=[BURGUNDY, TEAL], yfmt=ONE)),
    ('Trade Model', 'K52', dict(
        kind='doughnut', title='Export Destinations, FY2026A',
        data=(4, 29, 4, 35), cats=(1, 30, 1, 35), rows=False, points=6,
        colors=[NAVY, GOLD, TEAL, GREY, GREY, SKY], w=13.5, h=9.0)),
    ('Trade Model', 'R52', dict(
        kind='doughnut', title='Import Sources, FY2026A',
        data=(9, 29, 9, 35), cats=(6, 30, 6, 35), rows=False, points=6,
        colors=[BURGUNDY, STEEL, CHARCOAL, GREY, GREY, SKY], w=13.5, h=9.0)),

    # ---------------------------------------------------------- 20 ESG Model
    ('ESG Model', 'K7', dict(
        kind='combo', title='Industry CO2 Emissions and Emission Intensity',
        bar=(1, 19, 9, 19), line=(1, 20, 9, 20), cats=Y8, rows=True,
        bar_colors=[CHARCOAL], line_colors=[SAGE], yfmt=INT, y2fmt=TWO,
        yt='Mt CO2e', y2t='tCO2e per tonne')),
    ('ESG Model', 'K22', dict(
        kind='bar', title='EU CBAM Carbon Cost (Rs cr)',
        data=(1, 33, 8, 33), cats=(2, 30, 8, 30), rows=True, colors=[SAGE],
        yfmt=INT, legend=None, labels=True)),
    ('ESG Model', 'K37', dict(
        kind='bar', title='CBAM EBITDA Impact by Scenario, FY2033E (Rs cr)',
        data=(4, 85, 4, 89), cats=(1, 86, 1, 89), rows=False, colors=[BURGUNDY],
        yfmt=INT, legend=None, labels=True)),

    # ---------------------------------------------------------- 21 Scenario Manager
    ('Scenario Manager', 'K7', dict(
        kind='bar', title='FY2033E Revenue and EBITDA by Scenario (Rs cr)',
        data=(1, 28, 5, 29), cats=(2, 27, 5, 27), rows=True, colors=[NAVY, GOLD],
        yfmt=INT)),
    ('Scenario Manager', 'K22', dict(
        kind='bar', title='FY2033E Industry EBITDA Margin by Scenario',
        data=(2, 30, 5, 30), cats=(2, 27, 5, 27), rows=True, titles=False,
        colors=[NAVY, TEAL, GOLD, BURGUNDY], yfmt=PCT, legend=None, labels=True)),
    ('Scenario Manager', 'K37', dict(
        kind='bar', title='FY2033E Industry Enterprise Value by Scenario (Rs cr)',
        data=(2, 35, 5, 35), cats=(2, 27, 5, 27), rows=True, titles=False,
        colors=[NAVY, TEAL, GOLD, BURGUNDY], yfmt=INT, legend=None, labels=True)),
    ('Scenario Manager', 'K52', dict(
        kind='radar', title='Scenario Comparison - Spider',
        data=(2, 39, 5, 45), cats=(1, 40, 1, 45), rows=False,
        colors=[NAVY, TEAL, GOLD, BURGUNDY], yfmt=PCT, w=14.5, h=10.5)),

    # ---------------------------------------------------------- 22 Sensitivity Analysis
    ('Sensitivity Analysis', 'K7', dict(
        kind='hbar', title='Tornado: FY2033E EBITDA, 10% Adverse Move in Each Driver',
        data=(6, 96, 6, 106), cats=(1, 97, 1, 106), rows=False, colors=[BURGUNDY],
        yfmt=PCT, legend=None, labels=True, h=11.5)),
    ('Sensitivity Analysis', 'K31', dict(
        kind='line', title='FY2033E EBITDA against a Steel Price Shock (Rs cr)',
        data=(4, 18, 4, 25), cats=(2, 19, 2, 25), rows=False, colors=[NAVY],
        yfmt=INT, xt='Change in realisation', yt='EBITDA (Rs cr)', legend=None)),
    ('Sensitivity Analysis', 'K46', dict(
        kind='hbar', title='Downside and Upside Impact on FY2033E EBITDA (Rs cr)',
        data=(2, 55, 3, 63), cats=(1, 56, 1, 63), rows=False,
        colors=[BURGUNDY, TEAL], yfmt=INT, h=10.5)),
    ('Sensitivity Analysis', 'K64', dict(
        kind='line', title='Enterprise Value against WACC and Terminal Growth (Rs cr)',
        data=(1, 39, 6, 43), cats=(2, 38, 6, 38), rows=True, colors=PALETTE,
        yfmt=INT, xt='WACC', yt='Enterprise value (Rs cr)', h=9.6)),

    # ---------------------------------------------------------- 23 Comparable Valuation
    ('Comparable Valuation', 'K7', dict(
        kind='hbar', title='Football Field: Implied Enterprise Value by Method (Rs cr)',
        data=(2, 83, 2, 89), cats=(1, 84, 1, 89), rows=False,
        colors=[NAVY, STEEL, GOLD, GREY, TEAL, BURGUNDY], yfmt=INT, legend=None,
        labels=True, h=9.6)),
    ('Comparable Valuation', 'K27', dict(
        kind='bar', title='Industry Enterprise Value by Scenario, FY2033E (Rs cr)',
        data=(3, 93, 3, 97), cats=(1, 94, 1, 97), rows=False,
        colors=[NAVY, TEAL, GOLD, BURGUNDY], yfmt=INT, legend=None, labels=True)),
    ('Comparable Valuation', 'K43', dict(
        kind='bar', title='EV/EBITDA by Producer, FY2026A',
        data=(2, 38, 2, 43), cats=(1, 39, 1, 43), rows=False, colors=[STEEL],
        yfmt=X, legend=None, labels=True)),
    ('Comparable Valuation', 'K59', dict(
        kind='bar', title='Present Value of Unlevered Free Cash Flow (Rs cr)',
        data=(4, 131, 10, 131), cats=(4, 126, 10, 126), rows=True, titles=False,
        colors=[TEAL], yfmt=INT, legend=None, labels=True)),

    # ---------------------------------------------------------- 24 Industry Dashboard
    # The owner deliberately left rows 8 to 28 blank as a chart canvas. That canvas is
    # about 28.5 cm wide and 10.7 cm tall, which fits exactly three charts side by side.
    # Three further charts are placed underneath the Support & Audit table.
    ('Industry Dashboard', 'A8', dict(
        kind='combo', title='Industry Revenue and EBITDA Margin',
        bar=(3, 47, 10, 47), line=(3, 46, 10, 46), cats=(3, 30, 10, 30), rows=True,
        titles=False, bar_colors=[NAVY], line_colors=[GOLD], yfmt=INT, y2fmt=PCT,
        yt='Revenue (Rs cr)', y2t='Margin (%)', legend=None, w=8.9, h=10.2)),
    ('Industry Dashboard', 'B8', dict(
        kind='line', title='Apparent Consumption (Mt)',
        data=(3, 32, 10, 32), cats=(3, 30, 10, 30), rows=True, titles=False,
        colors=[STEEL], yfmt=INT, legend=None, w=8.9, h=10.2)),
    ('Industry Dashboard', 'G8', dict(
        kind='line', title='Capacity Utilisation',
        data=(3, 37, 10, 37), cats=(3, 30, 10, 30), rows=True, titles=False,
        colors=[BURGUNDY], yfmt=PCT, legend=None, w=8.9, h=10.2)),
]


# Placed below the Support & Audit table on 24 Industry Dashboard, where there is room
# for full-size charts. The anchor row is supplied by the builder at run time.
DASH_EXTRA = [
    dict(kind='line', title='Blended Realisation (Rs/t)',
         data=(3, 39, 10, 39), cats=(3, 30, 10, 30), rows=True, titles=False,
         colors=[NAVY], yfmt=INT, legend=None, w=18.0, h=9.0),
    dict(kind='bar', title='Unlevered Free Cash Flow (Rs cr)',
         data=(3, 52, 10, 52), cats=(3, 30, 10, 30), rows=True, titles=False,
         colors=[TEAL], yfmt=INT, legend=None, labels=True, w=18.0, h=9.0),
    dict(kind='combo', title='Net Debt and Leverage',
         bar=(3, 54, 10, 54), line=(3, 55, 10, 55), cats=(3, 30, 10, 30), rows=True,
         titles=False, bar_colors=[CHARCOAL], line_colors=[GOLD], yfmt=INT, y2fmt=X,
         yt='Net debt (Rs cr)', y2t='x EBITDA', legend=None, w=18.0, h=9.0),
]


def build(wb, dash_row=None):
    made = []
    for sheet, anchor, spec in SPECS:
        ws = wb[sheet]
        ch = make(ws, spec)
        ws.add_chart(ch, anchor)
        made.append((sheet, anchor, spec['title']))
    if dash_row:
        ws = wb['Industry Dashboard']
        for i, spec in enumerate(DASH_EXTRA):
            anchor = f'A{dash_row + i * 19}'
            ws.add_chart(make(ws, spec), anchor)
            made.append(('Industry Dashboard', anchor, spec['title']))
    return made
