#!/usr/bin/env python3
"""Generate March scorecard workbook and CSV templates (text-only source)."""
# Exec formatting v2

import csv
import zipfile
from xml.sax.saxutils import escape

XLSX_NAME = "March_Scorecard_Template.xlsx"
DAILY_CSV = "Daily_Inputs_Template.csv"
AR_CSV = "AR_Detail_Template.csv"


def c(ref, value=None, formula=None, style=None):
    attrs = [f'r="{ref}"']
    if style is not None:
        attrs.append(f's="{style}"')
    a = " " + " ".join(attrs)
    if formula is not None:
        return f'<c{a}><f>{escape(formula)}</f><v>0</v></c>'
    if value is None:
        return f'<c{a}/>'
    if isinstance(value, (int, float)):
        return f'<c{a}><v>{value}</v></c>'
    return f'<c{a} t="inlineStr"><is><t>{escape(str(value))}</t></is></c>'


def sheet_xml(rows, cols=None, cond=None, table_rids=None, freeze=None, page_setup=None):
    out = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    ]

    if freeze:
        x_split, y_split, top_left = freeze
        out.append('<sheetViews><sheetView workbookViewId="0">')
        out.append(f'<pane xSplit="{x_split}" ySplit="{y_split}" topLeftCell="{top_left}" activePane="bottomRight" state="frozen"/>')
        out.append('<selection pane="bottomRight"/>')
        out.append('</sheetView></sheetViews>')

    if cols:
        out.append('<cols>')
        for spec in cols:
            if len(spec) == 3:
                mi, ma, w = spec
                hidden = False
            else:
                mi, ma, w, hidden = spec
            out.append(f'<col min="{mi}" max="{ma}" width="{w}" customWidth="1"' + (' hidden="1"' if hidden else '') + '/>')
        out.append('</cols>')

    out.append('<sheetData>')
    for r in sorted(rows):
        out.append(f'<row r="{r}">{"".join(rows[r])}</row>')
    out.append('</sheetData>')

    if cond:
        out.extend(cond)

    if page_setup:
        out.append(page_setup)

    if table_rids:
        out.append(f'<tableParts count="{len(table_rids)}">')
        for rid in table_rids:
            out.append(f'<tablePart r:id="{rid}"/>')
        out.append('</tableParts>')

    out.append('</worksheet>')
    return ''.join(out)


# style ids
S_DEFAULT = 0
S_TITLE = 1
S_HEADER = 2
S_LABEL = 3
S_INPUT = 4
S_TEXT = 5
S_INT = 6
S_CUR = 7
S_PCT = 8
S_DATE = 9
S_WRAP = 10
S_TOTAL = 11


def build_assumptions():
    rows = {1: [c('A1', 'March Scorecard – Assumptions', style=S_TITLE)]}
    items = [
        (3, 'March Overhead', 560000, S_CUR),
        (4, 'March CM Target', 296000, S_CUR),
        (5, 'Working Days in March', 22, S_INT),
        (6, 'Field Headcount', 38, S_INT),
        (7, 'Hours per Day', 10, S_INT),
        (8, 'Capacity Hours', None, S_INT),
        (9, 'UMB/D&B Revenue Minimum', 165000, S_CUR),
        (10, 'UMB/D&B CM %', 0.65, S_PCT),
        (11, 'Sod Consumption Forecast (sq ft)', 921000, S_INT),
        (12, 'Sod Margin Delta', 0.00, S_PCT),
        (13, 'AR Days Plan', '', S_INT),
        (14, 'Warranty Unbillable Material Target', '', S_CUR),
        (15, 'Warranty Unbillable Labor Hours Target', '', S_INT),
    ]
    for r, label, val, vstyle in items:
        rows[r] = [c(f'A{r}', label, style=S_LABEL)]
        if r == 8:
            rows[r].append(c('B8', formula='B6*B5*B7', style=S_INPUT))
        else:
            rows[r].append(c(f'B{r}', val, style=S_INPUT if val == '' else vstyle))
    rows[17] = [c('A17', 'Notes', style=S_LABEL)]
    rows[18] = [c('A18', 'Sod Margin Delta allowed examples: 0.00, 0.05, 0.20', style=S_WRAP)]
    rows[19] = [c('A19', 'Headcount variance = projected average headcount - forecast headcount', style=S_WRAP)]
    return sheet_xml(rows, cols=[(1, 1, 48), (2, 2, 22)])


def build_forecast():
    rows = {1: [c('A1', 'March Forecast', style=S_TITLE)]}
    hdr = ['Category', 'March Revenue Forecast', 'CM %', 'CM $ (calculated)', 'Required Labor Hours', 'Notes']
    rows[3] = [c(f'{col}3', h, style=S_HEADER) for col, h in zip('ABCDEF', hdr)]

    for r, name in enumerate(['Production', 'LD', 'UMB/D&B'], start=4):
        rows[r] = [
            c(f'A{r}', name, style=S_TEXT),
            c(f'B{r}', 0, style=S_CUR),
            c(f'C{r}', 0, style=S_PCT),
            c(f'D{r}', formula=f'B{r}*C{r}', style=S_CUR),
            c(f'E{r}', 0, style=S_INT),
            c(f'F{r}', '', style=S_WRAP),
        ]
    rows[6][1] = c('B6', formula='Assumptions!B9', style=S_CUR)
    rows[6][2] = c('C6', formula='Assumptions!B10', style=S_PCT)

    rows[8] = [
        c('A8', 'Totals', style=S_TOTAL),
        c('B8', formula='SUM(B4:B6)', style=S_TOTAL),
        c('D8', formula='SUM(D4:D6)', style=S_TOTAL),
        c('E8', formula='SUM(E4:E6)', style=S_TOTAL),
    ]

    return sheet_xml(rows, cols=[(1, 1, 18), (2, 2, 20), (3, 3, 10), (4, 4, 16), (5, 5, 20), (6, 6, 26)], freeze=(0, 3, 'A4'), table_rids=['rId1'])


def build_daily_inputs():
    rows = {1: [c('A1', 'Daily Inputs (enter daily results)', style=S_TITLE)]}
    cols = [
        'Date', 'Revenue_Production', 'Revenue_LD', 'Revenue_UMB_D_B', 'CM_Production', 'CM_LD', 'CM_UMB_D_B',
        'Headcount_Field', 'Hours_Worked', 'Warranty_Unbillable_Material', 'Warranty_Unbillable_Labor_Hours'
    ]
    rows[3] = [c(f'{chr(64+i)}3', h, style=S_HEADER) for i, h in enumerate(cols, start=1)]
    for r in range(4, 36):
        rows[r] = [
            c(f'A{r}', '', style=S_DATE), c(f'B{r}', '', style=S_CUR), c(f'C{r}', '', style=S_CUR), c(f'D{r}', '', style=S_CUR),
            c(f'E{r}', '', style=S_CUR), c(f'F{r}', '', style=S_CUR), c(f'G{r}', '', style=S_CUR), c(f'H{r}', '', style=S_INT),
            c(f'I{r}', '', style=S_INT), c(f'J{r}', '', style=S_CUR), c(f'K{r}', '', style=S_INT)
        ]
        rows[r].append(c(f'M{r}', formula=f'IF(A{r}="","",IF(COUNTIF($A$4:A{r},A{r})=1,1,0))', style=S_INT))

    rows[2] = [c('N2', formula='SUM(M4:M35)', style=S_INT)]

    return sheet_xml(
        rows,
        cols=[(1, 1, 12), (2, 4, 16), (5, 7, 14), (8, 9, 12), (10, 11, 24), (13, 14, 12, True)],
        freeze=(1, 3, 'B4'),
        table_rids=['rId1'],
    )


def build_scorecard():
    rows = {
        1: [c('A1', 'March Scorecard (Executive View)', style=S_TITLE)],
        2: [c('A2', 'Revenue = Completed and Billed Only', style=S_LABEL)],
    }
    headers = ['Metric', 'March Forecast', 'MTD Actual', 'Avg per Day', 'Projected Month', 'Variance vs Forecast']
    rows[3] = [c(f'{ch}3', h, style=S_HEADER) for ch, h in zip('ABCDEF', headers)]

    metrics = [
        'Revenue D&B/UMB', 'Revenue LD', 'Revenue Production',
        'CM D&B/UMB', 'CM LD', 'CM Production',
        'Headcount', 'Labor Utilization %', 'AR Days to Pay (Plan vs Actual)',
        'Warranty Unbillable Material', 'Warranty Unbillable Labor'
    ]
    for r, m in enumerate(metrics, start=4):
        rows[r] = [c(f'A{r}', m, style=S_LABEL)]

    # Revenue + CM
    rows[4] += [c('B4', formula='Forecast!B6', style=S_CUR), c('C4', formula='SUM(Daily_Inputs!D4:D35)', style=S_CUR), c('D4', formula='IFERROR(C4/Daily_Inputs!N2,0)', style=S_CUR), c('E4', formula='D4*Assumptions!B5', style=S_CUR), c('F4', formula='E4-B4', style=S_CUR)]
    rows[5] += [c('B5', formula='Forecast!B5', style=S_CUR), c('C5', formula='SUM(Daily_Inputs!C4:C35)', style=S_CUR), c('D5', formula='IFERROR(C5/Daily_Inputs!N2,0)', style=S_CUR), c('E5', formula='D5*Assumptions!B5', style=S_CUR), c('F5', formula='E5-B5', style=S_CUR)]
    rows[6] += [c('B6', formula='Forecast!B4', style=S_CUR), c('C6', formula='SUM(Daily_Inputs!B4:B35)', style=S_CUR), c('D6', formula='IFERROR(C6/Daily_Inputs!N2,0)', style=S_CUR), c('E6', formula='D6*Assumptions!B5', style=S_CUR), c('F6', formula='E6-B6', style=S_CUR)]
    rows[7] += [c('B7', formula='Forecast!D6', style=S_CUR), c('C7', formula='SUM(Daily_Inputs!G4:G35)', style=S_CUR), c('D7', formula='IFERROR(C7/Daily_Inputs!N2,0)', style=S_CUR), c('E7', formula='D7*Assumptions!B5', style=S_CUR), c('F7', formula='E7-B7', style=S_CUR)]
    rows[8] += [c('B8', formula='Forecast!D5', style=S_CUR), c('C8', formula='SUM(Daily_Inputs!F4:F35)', style=S_CUR), c('D8', formula='IFERROR(C8/Daily_Inputs!N2,0)', style=S_CUR), c('E8', formula='D8*Assumptions!B5', style=S_CUR), c('F8', formula='E8-B8', style=S_CUR)]
    rows[9] += [c('B9', formula='Forecast!D4', style=S_CUR), c('C9', formula='SUM(Daily_Inputs!E4:E35)', style=S_CUR), c('D9', formula='IFERROR(C9/Daily_Inputs!N2,0)', style=S_CUR), c('E9', formula='D9*Assumptions!B5', style=S_CUR), c('F9', formula='E9-B9', style=S_CUR)]
    rows[10] += [c('B10', formula='Assumptions!B6', style=S_INT), c('C10', formula='IFERROR(AVERAGEIFS(Daily_Inputs!H4:H35,Daily_Inputs!A4:A35,"<>"),0)', style=S_INT), c('D10', formula='C10', style=S_INT), c('E10', formula='C10', style=S_INT), c('F10', formula='E10-B10', style=S_INT)]
    rows[11] += [c('B11', formula='IFERROR(Forecast!E8/Assumptions!B8,0)', style=S_PCT), c('C11', formula='IFERROR(SUM(Daily_Inputs!I4:I35)/(C10*Assumptions!B7*Daily_Inputs!N2),0)', style=S_PCT), c('D11', formula='C11', style=S_PCT), c('E11', formula='C11', style=S_PCT), c('F11', formula='E11-B11', style=S_PCT)]
    rows[12] += [c('B12', formula='Assumptions!B13', style=S_INT), c('C12', '', style=S_INPUT), c('D12', '', style=S_TEXT), c('E12', '', style=S_TEXT), c('F12', formula='IF(B12="","",IF(C12="","",C12-B12))', style=S_INT)]
    rows[13] += [c('B13', formula='Assumptions!B14', style=S_CUR), c('C13', formula='SUM(Daily_Inputs!J4:J35)', style=S_CUR), c('D13', formula='IFERROR(C13/Daily_Inputs!N2,0)', style=S_CUR), c('E13', formula='D13*Assumptions!B5', style=S_CUR), c('F13', formula='IF(B13="","",E13-B13)', style=S_CUR)]
    rows[14] += [c('B14', formula='Assumptions!B15', style=S_INT), c('C14', formula='SUM(Daily_Inputs!K4:K35)', style=S_INT), c('D14', formula='IFERROR(C14/Daily_Inputs!N2,0)', style=S_INT), c('E14', formula='D14*Assumptions!B5', style=S_INT), c('F14', formula='IF(B14="","",E14-B14)', style=S_INT)]

    cond = [
        '<conditionalFormatting sqref="F4:F11 F13:F14"><cfRule type="cellIs" dxfId="0" priority="1" operator="lessThan"><formula>0</formula></cfRule></conditionalFormatting>',
        '<conditionalFormatting sqref="C12"><cfRule type="expression" dxfId="0" priority="2"><formula>AND($B12&lt;&gt;"",$C12&lt;&gt;"",$C12&gt;$B12)</formula></cfRule></conditionalFormatting>',
        '<conditionalFormatting sqref="E13:E14"><cfRule type="expression" dxfId="0" priority="3"><formula>AND($B13&lt;&gt;"",$E13&gt;$B13)</formula></cfRule></conditionalFormatting>',
    ]

    scorecard_page = '<printOptions horizontalCentered="0" verticalCentered="0"/><pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.3" footer="0.3"/><pageSetup orientation="landscape" fitToWidth="1" fitToHeight="0"/>'
    return sheet_xml(rows, cols=[(1, 1, 38), (2, 6, 18)], cond=cond, freeze=(1, 3, 'B4'), page_setup=scorecard_page)


def build_capacity():
    rows = {1: [c('A1', 'Capacity Overview', style=S_TITLE)]}
    rows[3] = [c('A3', 'Available Capacity Hours', style=S_LABEL), c('B3', formula='Assumptions!B8', style=S_INT)]
    rows[4] = [c('A4', 'Required Hours', style=S_LABEL), c('B4', formula='Forecast!E8', style=S_INT)]
    rows[5] = [c('A5', 'Actual Hours Worked', style=S_LABEL), c('B5', formula='SUM(Daily_Inputs!I4:I35)', style=S_INT)]
    rows[6] = [c('A6', 'Remaining Capacity', style=S_LABEL), c('B6', formula='B3-B5', style=S_INT)]
    rows[7] = [c('A7', 'Utilization %', style=S_LABEL), c('B7', formula='IFERROR(B5/B3,0)', style=S_PCT)]
    cond = ['<conditionalFormatting sqref="B7"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0.95</formula></cfRule></conditionalFormatting>']
    return sheet_xml(rows, cols=[(1, 1, 32), (2, 2, 20)], cond=cond)


def build_cashflow():
    rows = {1: [c('A1', 'Weekly Cashflow - March', style=S_TITLE)]}
    hdr = ['Week', 'Beginning Cash', 'Revenue Collected', 'Overhead Allocation', 'Payroll Placeholder', 'Equipment Proceeds', 'Bowman Cash', 'Ending Cash']
    rows[3] = [c(f'{chr(64+i)}3', h, style=S_HEADER) for i, h in enumerate(hdr, start=1)]
    for i, r in enumerate(range(4, 8), start=1):
        rows[r] = [c(f'A{r}', f'Week {i}', style=S_TEXT)]
        rows[r].append(c(f'B{r}', 0 if r == 4 else None, formula=None if r == 4 else f'H{r-1}', style=S_CUR))
        rows[r] += [
            c(f'C{r}', formula='(Scorecard!E4+Scorecard!E5+Scorecard!E6)/4', style=S_CUR),
            c(f'D{r}', formula='Assumptions!B3/4', style=S_CUR),
            c(f'E{r}', 0, style=S_CUR),
            c(f'F{r}', 0, style=S_CUR),
            c(f'G{r}', 0, style=S_CUR),
            c(f'H{r}', formula=f'B{r}+C{r}-D{r}-E{r}+F{r}+G{r}', style=S_CUR),
        ]
    rows[10] = [c('A10', 'Scenario Placeholders', style=S_LABEL)]
    rows[11] = [c('A11', 'Base Case', style=S_LABEL)]
    rows[12] = [c('A12', 'Conservative Case', style=S_LABEL)]
    rows[13] = [c('A13', 'Stress Case', style=S_LABEL)]
    return sheet_xml(rows, cols=[(1, 1, 14), (2, 8, 18)], table_rids=['rId1'])


def build_xlsx():
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet5.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet6.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
<Override PartName="/xl/tables/table2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
<Override PartName="/xl/tables/table3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
</Types>'''

    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

    workbook_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Assumptions" sheetId="1" r:id="rId1"/>
<sheet name="Forecast" sheetId="2" r:id="rId2"/>
<sheet name="Daily_Inputs" sheetId="3" r:id="rId3"/>
<sheet name="Scorecard" sheetId="4" r:id="rId4"/>
<sheet name="Capacity" sheetId="5" r:id="rId5"/>
<sheet name="Cashflow" sheetId="6" r:id="rId6"/>
</sheets>
<definedNames>
<definedName name="_xlnm.Print_Area" localSheetId="3">Scorecard!$A$1:$F$14</definedName>
<definedName name="_xlnm.Print_Titles" localSheetId="3">Scorecard!$3:$3</definedName>
</definedNames>
</workbook>'''

    workbook_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet4.xml"/>
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet5.xml"/>
<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet6.xml"/>
<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="$#,##0"/><numFmt numFmtId="165" formatCode="0.0%"/></numFmts>
<fonts count="3"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="12"/><name val="Calibri"/><family val="2"/></font></fonts>
<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDCE6F1"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="12">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
<xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="14" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill><font><color rgb="FF9C0006"/></font></dxf></dxfs>
<tableStyles count="0" defaultTableStyle="TableStyleLight9" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>'''

    sheet2_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table2.xml"/>
</Relationships>'''

    sheet3_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>'''

    sheet6_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table3.xml"/>
</Relationships>'''

    table1_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="tblDailyInputs" displayName="tblDailyInputs" ref="A3:K35" totalsRowShown="0">
<autoFilter ref="A3:K35"/>
<tableColumns count="11">
<tableColumn id="1" name="Date"/><tableColumn id="2" name="Revenue_Production"/><tableColumn id="3" name="Revenue_LD"/><tableColumn id="4" name="Revenue_UMB_D_B"/><tableColumn id="5" name="CM_Production"/><tableColumn id="6" name="CM_LD"/><tableColumn id="7" name="CM_UMB_D_B"/><tableColumn id="8" name="Headcount_Field"/><tableColumn id="9" name="Hours_Worked"/><tableColumn id="10" name="Warranty_Unbillable_Material"/><tableColumn id="11" name="Warranty_Unbillable_Labor_Hours"/>
</tableColumns>
<tableStyleInfo name="TableStyleLight9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>'''

    table2_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="2" name="tblForecast" displayName="tblForecast" ref="A3:F6" totalsRowShown="0">
<autoFilter ref="A3:F6"/>
<tableColumns count="6">
<tableColumn id="1" name="Category"/><tableColumn id="2" name="March Revenue Forecast"/><tableColumn id="3" name="CM %"/><tableColumn id="4" name="CM $ (calculated)"/><tableColumn id="5" name="Required Labor Hours"/><tableColumn id="6" name="Notes"/>
</tableColumns>
<tableStyleInfo name="TableStyleLight9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>'''

    table3_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="3" name="tblCashflow" displayName="tblCashflow" ref="A3:H7" totalsRowShown="0">
<autoFilter ref="A3:H7"/>
<tableColumns count="8">
<tableColumn id="1" name="Week"/><tableColumn id="2" name="Beginning Cash"/><tableColumn id="3" name="Revenue Collected"/><tableColumn id="4" name="Overhead Allocation"/><tableColumn id="5" name="Payroll Placeholder"/><tableColumn id="6" name="Equipment Proceeds"/><tableColumn id="7" name="Bowman Cash"/><tableColumn id="8" name="Ending Cash"/>
</tableColumns>
<tableStyleInfo name="TableStyleLight9" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>'''

    with zipfile.ZipFile(XLSX_NAME, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', content_types)
        z.writestr('_rels/.rels', root_rels)
        z.writestr('xl/workbook.xml', workbook_xml)
        z.writestr('xl/_rels/workbook.xml.rels', workbook_rels)
        z.writestr('xl/styles.xml', styles_xml)
        z.writestr('xl/worksheets/sheet1.xml', build_assumptions())
        z.writestr('xl/worksheets/sheet2.xml', build_forecast())
        z.writestr('xl/worksheets/sheet3.xml', build_daily_inputs())
        z.writestr('xl/worksheets/sheet4.xml', build_scorecard())
        z.writestr('xl/worksheets/sheet5.xml', build_capacity())
        z.writestr('xl/worksheets/sheet6.xml', build_cashflow())
        z.writestr('xl/worksheets/_rels/sheet2.xml.rels', sheet2_rels)
        z.writestr('xl/worksheets/_rels/sheet3.xml.rels', sheet3_rels)
        z.writestr('xl/worksheets/_rels/sheet6.xml.rels', sheet6_rels)
        z.writestr('xl/tables/table1.xml', table1_xml)
        z.writestr('xl/tables/table2.xml', table2_xml)
        z.writestr('xl/tables/table3.xml', table3_xml)


def build_csvs():
    daily_headers = [
        'Date', 'Revenue_Production', 'Revenue_LD', 'Revenue_UMB_D_B', 'CM_Production', 'CM_LD', 'CM_UMB_D_B',
        'Headcount_Field', 'Hours_Worked', 'Warranty_Unbillable_Material', 'Warranty_Unbillable_Labor_Hours'
    ]
    with open(DAILY_CSV, 'w', newline='', encoding='utf-8') as f:
        csv.writer(f).writerow(daily_headers)

    ar_headers = ['Invoice_Number', 'Customer', 'Invoice_Date', 'Due_Date', 'Amount', 'Amount_Collected', 'Balance_Remaining', 'Days_Outstanding', 'Status', 'Notes']
    with open(AR_CSV, 'w', newline='', encoding='utf-8') as f:
        csv.writer(f).writerow(ar_headers)


if __name__ == '__main__':
    build_xlsx()
    build_csvs()
    print(f'Generated {XLSX_NAME}, {DAILY_CSV}, and {AR_CSV}')
