#!/usr/bin/env python3
"""Generate March scorecard workbook and CSV templates."""

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
    attrs_txt = " " + " ".join(attrs)

    if formula is not None:
        return f'<c{attrs_txt}><f>{escape(formula)}</f><v>0</v></c>'
    if value is None:
        return f'<c{attrs_txt}/>'
    if isinstance(value, (int, float)):
        return f'<c{attrs_txt}><v>{value}</v></c>'
    return f'<c{attrs_txt} t="inlineStr"><is><t>{escape(str(value))}</t></is></c>'


def sheet_xml(rows, cols=None, cond=None, table_part=False):
    out = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    ]
    if cols:
        out.append('<cols>')
        for spec in cols:
            if len(spec) == 3:
                min_c, max_c, width = spec
                hidden = False
            else:
                min_c, max_c, width, hidden = spec
            hidden_attr = ' hidden="1"' if hidden else ''
            out.append(f'<col min="{min_c}" max="{max_c}" width="{width}" customWidth="1"{hidden_attr}/>' )
        out.append('</cols>')

    out.append('<sheetData>')
    for r in sorted(rows.keys()):
        out.append(f'<row r="{r}">{"".join(rows[r])}</row>')
    out.append('</sheetData>')

    if cond:
        out.extend(cond)
    if table_part:
        out.append('<tableParts count="1"><tablePart r:id="rId1"/></tableParts>')
    out.append('</worksheet>')
    return ''.join(out)


def build_assumptions():
    rows = {1: [c('A1', 'March Scorecard Assumptions', style=1)]}
    items = [
        (3, 'March Overhead', 560000),
        (4, 'March CM Target', 296000),
        (5, 'Working Days in March', 22),
        (6, 'Field Headcount', 38),
        (7, 'Hours per Day', 10),
        (8, 'Capacity Hours', None),
        (9, 'UMB/D&B Revenue Minimum', 165000),
        (10, 'UMB/D&B CM %', 0.65),
        (11, 'Sod Consumption Forecast (sq ft)', 921000),
        (12, 'Sod Margin Delta', 0.00),
        (13, 'AR Days Plan', ''),
        (14, 'Warranty Unbillable Material Target', ''),
        (15, 'Warranty Unbillable Labor Hours Target', ''),
    ]
    for row, label, val in items:
        rows[row] = [c(f'A{row}', label, style=1)]
        if row == 8:
            rows[row].append(c('B8', formula='B6*B5*B7'))
        else:
            rows[row].append(c(f'B{row}', val))
    rows[17] = [c('A17', 'Notes', style=1)]
    rows[18] = [c('A18', 'Sod Margin Delta allowed examples: 0.00, 0.05, 0.20')]
    rows[19] = [c('A19', 'Headcount variance = Projected average headcount - forecast headcount')]
    return sheet_xml(rows, cols=[(1, 1, 46), (2, 2, 24)])


def build_forecast():
    rows = {1: [c('A1', 'March Forecast', style=1)]}
    hdr = ['Category', 'March Revenue Forecast', 'CM %', 'CM $ (calculated)', 'Required Labor Hours', 'Notes']
    rows[3] = [c(f'{col}3', h, style=2) for col, h in zip('ABCDEF', hdr)]
    for r, name in enumerate(['Production', 'LD', 'UMB/D&B'], start=4):
        rows[r] = [
            c(f'A{r}', name),
            c(f'B{r}', 0),
            c(f'C{r}', 0),
            c(f'D{r}', formula=f'B{r}*C{r}'),
            c(f'E{r}', 0),
            c(f'F{r}', ''),
        ]
    rows[6][1] = c('B6', formula='Assumptions!B9')
    rows[6][2] = c('C6', formula='Assumptions!B10')
    rows[8] = [c('A8', 'Totals', style=1), c('B8', formula='SUM(B4:B6)'), c('D8', formula='SUM(D4:D6)'), c('E8', formula='SUM(E4:E6)')]
    return sheet_xml(rows, cols=[(1, 1, 20), (2, 2, 24), (3, 3, 12), (4, 4, 18), (5, 5, 22), (6, 6, 24)])


def build_daily_inputs():
    rows = {1: [c('A1', 'Daily Inputs', style=1)]}
    cols = [
        'Date', 'Revenue_Production', 'Revenue_LD', 'Revenue_UMB_D_B', 'CM_Production', 'CM_LD', 'CM_UMB_D_B',
        'Headcount_Field', 'Hours_Worked', 'Warranty_Unbillable_Material', 'Warranty_Unbillable_Labor_Hours', 'Day_Flag'
    ]
    rows[3] = [c(f'{chr(64+i)}3', h, style=2) for i, h in enumerate(cols, start=1)]
    for r in range(4, 36):  # 32 prepared rows
        rows[r] = [c(f'{chr(64+i)}{r}', '') for i in range(1, 12)]
        rows[r].append(c(f'L{r}', formula=f'IF(A{r}="","",IF(COUNTIF($A$4:A{r},A{r})=1,1,0))'))

    rows[2] = [c('N2', formula='SUM(L4:L35)')]
    return sheet_xml(
        rows,
        cols=[(1, 1, 14), (2, 4, 18), (5, 7, 14), (8, 9, 14), (10, 11, 28), (12, 12, 10, True), (14, 14, 12, True)],
        table_part=True,
    )


def build_scorecard():
    rows = {1: [c('A1', 'Revenue = Completed and Billed Only', style=1)]}
    headers = ['Metric', 'March Forecast', 'MTD Actual', 'Avg per Day', 'Projected Month', 'Variance vs Forecast']
    rows[3] = [c(f'{ch}3', h, style=2) for ch, h in zip('ABCDEF', headers)]

    metrics = [
        'Revenue D&B/UMB', 'Revenue LD', 'Revenue Production',
        'CM D&B/UMB', 'CM LD', 'CM Production',
        'Headcount', 'Labor Utilization %', 'AR Days to Pay (Plan vs Actual)',
        'Warranty Unbillable Material', 'Warranty Unbillable Labor'
    ]
    for r, m in enumerate(metrics, start=4):
        rows[r] = [c(f'A{r}', m)]

    # Revenue + CM rows
    rows[4] += [c('B4', formula='Forecast!B6'), c('C4', formula='SUM(Daily_Inputs!D4:D35)'), c('D4', formula='IFERROR(C4/Daily_Inputs!N2,0)'), c('E4', formula='D4*Assumptions!B5'), c('F4', formula='E4-B4')]
    rows[5] += [c('B5', formula='Forecast!B5'), c('C5', formula='SUM(Daily_Inputs!C4:C35)'), c('D5', formula='IFERROR(C5/Daily_Inputs!N2,0)'), c('E5', formula='D5*Assumptions!B5'), c('F5', formula='E5-B5')]
    rows[6] += [c('B6', formula='Forecast!B4'), c('C6', formula='SUM(Daily_Inputs!B4:B35)'), c('D6', formula='IFERROR(C6/Daily_Inputs!N2,0)'), c('E6', formula='D6*Assumptions!B5'), c('F6', formula='E6-B6')]

    rows[7] += [c('B7', formula='Forecast!D6'), c('C7', formula='SUM(Daily_Inputs!G4:G35)'), c('D7', formula='IFERROR(C7/Daily_Inputs!N2,0)'), c('E7', formula='D7*Assumptions!B5'), c('F7', formula='E7-B7')]
    rows[8] += [c('B8', formula='Forecast!D5'), c('C8', formula='SUM(Daily_Inputs!F4:F35)'), c('D8', formula='IFERROR(C8/Daily_Inputs!N2,0)'), c('E8', formula='D8*Assumptions!B5'), c('F8', formula='E8-B8')]
    rows[9] += [c('B9', formula='Forecast!D4'), c('C9', formula='SUM(Daily_Inputs!E4:E35)'), c('D9', formula='IFERROR(C9/Daily_Inputs!N2,0)'), c('E9', formula='D9*Assumptions!B5'), c('F9', formula='E9-B9')]

    # Headcount
    rows[10] += [c('B10', formula='Assumptions!B6'), c('C10', formula='IFERROR(AVERAGEIFS(Daily_Inputs!H4:H35,Daily_Inputs!A4:A35,"<>"),0)'), c('D10', formula='C10'), c('E10', formula='C10'), c('F10', formula='E10-B10')]

    # Labor utilization
    rows[11] += [c('B11', formula='IFERROR(Forecast!E8/Assumptions!B8,0)'), c('C11', formula='IFERROR(SUM(Daily_Inputs!I4:I35)/(C10*Assumptions!B7*Daily_Inputs!N2),0)'), c('D11', formula='C11'), c('E11', formula='C11'), c('F11', formula='E11-B11')]

    # AR row (non-cumulative)
    rows[12] += [c('B12', formula='Assumptions!B13'), c('C12', ''), c('D12', ''), c('E12', ''), c('F12', '')]

    # Warranty rows using new targets
    rows[13] += [c('B13', formula='Assumptions!B14'), c('C13', formula='SUM(Daily_Inputs!J4:J35)'), c('D13', formula='IFERROR(C13/Daily_Inputs!N2,0)'), c('E13', formula='D13*Assumptions!B5'), c('F13', formula='E13-B13')]
    rows[14] += [c('B14', formula='Assumptions!B15'), c('C14', formula='SUM(Daily_Inputs!K4:K35)'), c('D14', formula='IFERROR(C14/Daily_Inputs!N2,0)'), c('E14', formula='D14*Assumptions!B5'), c('F14', formula='E14-B14')]

    cond = [
        '<conditionalFormatting sqref="F4:F11 F13:F14"><cfRule type="cellIs" dxfId="0" priority="1" operator="lessThan"><formula>0</formula></cfRule></conditionalFormatting>',
        '<conditionalFormatting sqref="C12"><cfRule type="expression" dxfId="0" priority="2"><formula>AND(C12&lt;&gt;"",B12&lt;&gt;"",C12&gt;B12)</formula></cfRule></conditionalFormatting>',
        '<conditionalFormatting sqref="E13:E14"><cfRule type="expression" dxfId="0" priority="3"><formula>AND(B13&lt;&gt;"",E13&gt;B13)</formula></cfRule></conditionalFormatting>',
    ]
    return sheet_xml(rows, cols=[(1, 1, 38), (2, 6, 18)], cond=cond)


def build_capacity():
    rows = {1: [c('A1', 'Capacity Overview', style=1)]}
    rows[3] = [c('A3', 'Available Capacity Hours', style=1), c('B3', formula='Assumptions!B8')]
    rows[4] = [c('A4', 'Required Hours', style=1), c('B4', formula='Forecast!E8')]
    rows[5] = [c('A5', 'Actual Hours Worked', style=1), c('B5', formula='SUM(Daily_Inputs!I4:I35)')]
    rows[6] = [c('A6', 'Remaining Capacity', style=1), c('B6', formula='B3-B5')]
    rows[7] = [c('A7', 'Utilization %', style=1), c('B7', formula='IFERROR(B5/B3,0)')]
    cond = ['<conditionalFormatting sqref="B7"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0.95</formula></cfRule></conditionalFormatting>']
    return sheet_xml(rows, cols=[(1, 1, 32), (2, 2, 20)], cond=cond)


def build_cashflow():
    rows = {1: [c('A1', 'Weekly Cashflow - March', style=1)]}
    hdr = ['Week', 'Beginning Cash', 'Revenue Collected', 'Overhead Allocation', 'Payroll Placeholder', 'Equipment Proceeds', 'Bowman Cash', 'Ending Cash']
    rows[3] = [c(f'{chr(64+i)}3', h, style=2) for i, h in enumerate(hdr, start=1)]
    for i, r in enumerate(range(4, 8), start=1):
        rows[r] = [c(f'A{r}', f'Week {i}')]
        rows[r].append(c(f'B{r}', 0 if r == 4 else None, formula=None if r == 4 else f'H{r-1}'))
        rows[r] += [
            c(f'C{r}', formula='(Scorecard!E4+Scorecard!E5+Scorecard!E6)/4'),
            c(f'D{r}', formula='Assumptions!B3/4'),
            c(f'E{r}', 0),
            c(f'F{r}', 0),
            c(f'G{r}', 0),
            c(f'H{r}', formula=f'B{r}+C{r}-D{r}-E{r}+F{r}+G{r}')
        ]
    rows[10] = [c('A10', 'Scenario Placeholders', style=1)]
    rows[11] = [c('A11', 'Base Case')]
    rows[12] = [c('A12', 'Conservative Case')]
    rows[13] = [c('A13', 'Stress Case')]
    return sheet_xml(rows, cols=[(1, 1, 14), (2, 8, 18)])


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
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>
<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill><font><color rgb="FF9C0006"/></font></dxf></dxfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

    sheet3_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>'''

    table_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="tblDailyInputs" displayName="tblDailyInputs" ref="A3:L35" totalsRowShown="0">
<autoFilter ref="A3:L35"/>
<tableColumns count="12">
<tableColumn id="1" name="Date"/><tableColumn id="2" name="Revenue_Production"/><tableColumn id="3" name="Revenue_LD"/><tableColumn id="4" name="Revenue_UMB_D_B"/><tableColumn id="5" name="CM_Production"/><tableColumn id="6" name="CM_LD"/><tableColumn id="7" name="CM_UMB_D_B"/><tableColumn id="8" name="Headcount_Field"/><tableColumn id="9" name="Hours_Worked"/><tableColumn id="10" name="Warranty_Unbillable_Material"/><tableColumn id="11" name="Warranty_Unbillable_Labor_Hours"/><tableColumn id="12" name="Day_Flag"/>
</tableColumns>
<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
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
        z.writestr('xl/worksheets/_rels/sheet3.xml.rels', sheet3_rels)
        z.writestr('xl/tables/table1.xml', table_xml)


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
