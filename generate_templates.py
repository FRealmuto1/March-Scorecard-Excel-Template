#!/usr/bin/env python3
"""Generate March scorecard workbook and CSV templates.

This keeps source in text form (no binary files committed) and produces:
- March_Scorecard_Template.xlsx
- Daily_Inputs_Template.csv
- AR_Detail_Template.csv
"""

import csv
import zipfile
from xml.sax.saxutils import escape


XLSX_NAME = "March_Scorecard_Template.xlsx"
DAILY_CSV = "Daily_Inputs_Template.csv"
AR_CSV = "AR_Detail_Template.csv"


def make_cell(ref, value=None, formula=None, style=None):
    attrs = f' r="{ref}"'
    if style is not None:
        attrs += f' s="{style}"'

    if formula is not None:
        calc_value = 0 if value is None else value
        return f'<c{attrs}><f>{escape(str(formula))}</f><v>{calc_value}</v></c>'

    if isinstance(value, (int, float)):
        return f'<c{attrs}><v>{value}</v></c>'

    if value is None:
        return f'<c{attrs}/>'

    return f'<c{attrs} t="inlineStr"><is><t>{escape(str(value))}</t></is></c>'


def build_sheet_xml(rows, cols=None, conditional_blocks=None, include_table_part=False):
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    ]

    if cols:
        parts.append('<cols>')
        for min_col, max_col, width in cols:
            parts.append(
                f'<col min="{min_col}" max="{max_col}" width="{width}" customWidth="1"/>'
            )
        parts.append('</cols>')

    parts.append('<sheetData>')
    for row_num in sorted(rows):
        row_cells = ''.join(rows[row_num])
        parts.append(f'<row r="{row_num}">{row_cells}</row>')
    parts.append('</sheetData>')

    if conditional_blocks:
        parts.extend(conditional_blocks)

    if include_table_part:
        parts.append('<tableParts count="1"><tablePart r:id="rId1"/></tableParts>')

    parts.append('</worksheet>')
    return ''.join(parts)


def build_assumptions_sheet():
    rows = {1: [make_cell("A1", "March Scorecard Assumptions", style=1)]}
    assumptions = [
        (3, "March Overhead", 560000),
        (4, "March CM Target", 296000),
        (5, "Working Days in March", 22),
        (6, "Field Headcount", 38),
        (7, "Hours per Day", 10),
        (8, "Capacity Hours", None),
        (9, "UMB/D&B Revenue Minimum", 165000),
        (10, "UMB/D&B CM %", 0.65),
        (11, "Sod Consumption Forecast (sq ft)", 921000),
        (12, "Sod Margin Delta", 0.00),
        (13, "AR Days Plan", 45),
    ]

    for row_num, label, value in assumptions:
        rows.setdefault(row_num, []).append(make_cell(f"A{row_num}", label, style=1))
        if row_num == 8:
            rows[row_num].append(make_cell("B8", formula="B6*B5*B7"))
        else:
            rows[row_num].append(make_cell(f"B{row_num}", value))

    rows[14] = [make_cell("A14", "Allowed Sod Margin Delta examples: 0.00, 0.05, 0.20")]
    return build_sheet_xml(rows, cols=[(1, 1, 42), (2, 2, 24)])


def build_forecast_sheet():
    rows = {1: [make_cell("A1", "March Forecast", style=1)]}
    headers = ["Category", "March Revenue Forecast", "CM %", "CM $ (calculated)", "Required Labor Hours", "Notes"]
    rows[3] = [make_cell(f"{col}3", title, style=2) for col, title in zip("ABCDEF", headers)]

    for row_num, category in enumerate(["Production", "LD", "UMB/D&B"], start=4):
        rows[row_num] = [
            make_cell(f"A{row_num}", category),
            make_cell(f"B{row_num}", 0),
            make_cell(f"C{row_num}", 0),
            make_cell(f"D{row_num}", formula=f"B{row_num}*C{row_num}"),
            make_cell(f"E{row_num}", 0),
            make_cell(f"F{row_num}", ""),
        ]

    rows[6][1] = make_cell("B6", formula="Assumptions!B9")
    rows[6][2] = make_cell("C6", formula="Assumptions!B10")

    rows[8] = [
        make_cell("A8", "Totals", style=1),
        make_cell("B8", formula="SUM(B4:B6)"),
        make_cell("D8", formula="SUM(D4:D6)"),
        make_cell("E8", formula="SUM(E4:E6)"),
    ]

    return build_sheet_xml(rows, cols=[(1, 1, 20), (2, 2, 24), (3, 3, 10), (4, 4, 18), (5, 5, 22), (6, 6, 24)])


def build_daily_inputs_sheet():
    rows = {1: [make_cell("A1", "Daily Inputs", style=1)]}
    columns = [
        "Date",
        "Revenue_Production",
        "Revenue_LD",
        "Revenue_UMB_D_B",
        "CM_Production",
        "CM_LD",
        "CM_UMB_D_B",
        "Headcount_Field",
        "Hours_Worked",
        "Warranty_Unbillable_Material",
        "Warranty_Unbillable_Labor_Hours",
    ]
    rows[3] = [make_cell(f"{chr(64 + idx)}3", name, style=2) for idx, name in enumerate(columns, start=1)]

    for row_num in range(4, 35):
        rows[row_num] = [make_cell(f"{chr(64 + idx)}{row_num}", "") for idx in range(1, 12)]

    return build_sheet_xml(
        rows,
        cols=[(1, 1, 14), (2, 4, 18), (5, 7, 14), (8, 9, 14), (10, 11, 28)],
        include_table_part=True,
    )


def build_scorecard_sheet():
    rows = {1: [make_cell("A1", "Revenue = Completed and Billed Only", style=1)]}
    headers = ["Metric", "March Forecast", "MTD Actual", "Avg per Day", "Projected Month", "Variance vs Forecast", "AR Actual"]
    rows[3] = [make_cell(f"{chr(64 + idx)}3", h, style=2) for idx, h in enumerate(headers, start=1)]

    metrics = [
        "Revenue D&B/UMB",
        "Revenue LD",
        "Revenue Production",
        "CM D&B/UMB",
        "CM LD",
        "CM Production",
        "Headcount",
        "Labor Utilization %",
        "AR Days to Pay (Plan vs Actual)",
        "Warranty Unbillable Material",
        "Warranty Unbillable Labor",
    ]
    for row_num, metric in enumerate(metrics, start=4):
        rows[row_num] = [make_cell(f"A{row_num}", metric)]

    forecast_formulas = [
        "Forecast!B6", "Forecast!B5", "Forecast!B4", "Forecast!D6", "Forecast!D5", "Forecast!D4",
        "Assumptions!B6", "Capacity!B7", "Assumptions!B13", "0", "0"
    ]
    actual_formulas = [
        "SUM(Daily_Inputs!D:D)", "SUM(Daily_Inputs!C:C)", "SUM(Daily_Inputs!B:B)",
        "SUM(Daily_Inputs!G:G)", "SUM(Daily_Inputs!F:F)", "SUM(Daily_Inputs!E:E)",
        "IFERROR(AVERAGE(Daily_Inputs!H:H),0)", "Capacity!B7", "0",
        "SUM(Daily_Inputs!J:J)", "SUM(Daily_Inputs!K:K)"
    ]

    for offset, row_num in enumerate(range(4, 15)):
        rows[row_num].append(make_cell(f"B{row_num}", formula=forecast_formulas[offset]))
        rows[row_num].append(make_cell(f"C{row_num}", formula=actual_formulas[offset]))
        rows[row_num].append(make_cell(f"D{row_num}", formula=f"IFERROR(C{row_num}/$H$4,0)"))
        rows[row_num].append(make_cell(f"E{row_num}", formula=f"D{row_num}*Assumptions!B5"))
        rows[row_num].append(make_cell(f"F{row_num}", formula=f"E{row_num}-B{row_num}"))
        if row_num == 12:
            rows[row_num].append(make_cell("G12", ""))

    rows[3].append(make_cell("H3", "Distinct Days"))
    rows[4].append(make_cell("H4", formula='COUNTA(UNIQUE(FILTER(Daily_Inputs!A:A,Daily_Inputs!A:A<>"")))'))

    conditional_blocks = [
        '<conditionalFormatting sqref="F4:F14"><cfRule type="cellIs" dxfId="0" priority="1" operator="lessThan"><formula>0</formula></cfRule></conditionalFormatting>',
        '<conditionalFormatting sqref="G12"><cfRule type="expression" dxfId="0" priority="2"><formula>G12&gt;B12</formula></cfRule></conditionalFormatting>',
        '<conditionalFormatting sqref="F13:F14"><cfRule type="cellIs" dxfId="0" priority="3" operator="lessThan"><formula>0</formula></cfRule></conditionalFormatting>',
    ]

    return build_sheet_xml(rows, cols=[(1, 1, 36), (2, 7, 16), (8, 8, 12)], conditional_blocks=conditional_blocks)


def build_capacity_sheet():
    rows = {1: [make_cell("A1", "Capacity Overview", style=1)]}
    entries = [
        (3, "Available Capacity Hours", "Assumptions!B8"),
        (4, "Required Hours", "Forecast!E8"),
        (5, "Actual Hours Worked", "SUM(Daily_Inputs!I:I)"),
        (6, "Remaining Capacity", "B3-B5"),
        (7, "Utilization %", "IFERROR(B5/B3,0)"),
    ]
    for row_num, label, formula in entries:
        rows[row_num] = [make_cell(f"A{row_num}", label, style=1), make_cell(f"B{row_num}", formula=formula)]

    conditional_blocks = [
        '<conditionalFormatting sqref="B7"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>0.95</formula></cfRule></conditionalFormatting>'
    ]
    return build_sheet_xml(rows, cols=[(1, 1, 32), (2, 2, 20)], conditional_blocks=conditional_blocks)


def build_cashflow_sheet():
    rows = {1: [make_cell("A1", "Weekly Cashflow - March", style=1)]}
    headers = [
        "Week", "Beginning Cash", "Revenue Collected", "Overhead Allocation",
        "Payroll Placeholder", "Equipment Proceeds", "Bowman Cash", "Ending Cash"
    ]
    rows[3] = [make_cell(f"{chr(64 + idx)}3", h, style=2) for idx, h in enumerate(headers, start=1)]

    for week_idx, row_num in enumerate(range(4, 8), start=1):
        row = [make_cell(f"A{row_num}", f"Week {week_idx}")]
        if row_num == 4:
            row.append(make_cell("B4", 0))
        else:
            row.append(make_cell(f"B{row_num}", formula=f"H{row_num - 1}"))

        row.extend([
            make_cell(f"C{row_num}", formula="(Scorecard!E4+Scorecard!E5+Scorecard!E6)/4"),
            make_cell(f"D{row_num}", formula="Assumptions!B3/4"),
            make_cell(f"E{row_num}", 0),
            make_cell(f"F{row_num}", 0),
            make_cell(f"G{row_num}", 0),
            make_cell(f"H{row_num}", formula=f"B{row_num}+C{row_num}-D{row_num}-E{row_num}+F{row_num}+G{row_num}"),
        ])
        rows[row_num] = row

    rows[10] = [make_cell("A10", "Scenario Placeholders", style=1)]
    rows[11] = [make_cell("A11", "Base Case")]
    rows[12] = [make_cell("A12", "Conservative Case")]
    rows[13] = [make_cell("A13", "Stress Case")]

    return build_sheet_xml(rows, cols=[(1, 1, 14), (2, 8, 18)])


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
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="tblDailyInputs" displayName="tblDailyInputs" ref="A3:K34" totalsRowShown="0">
<autoFilter ref="A3:K34"/>
<tableColumns count="11">
<tableColumn id="1" name="Date"/><tableColumn id="2" name="Revenue_Production"/><tableColumn id="3" name="Revenue_LD"/><tableColumn id="4" name="Revenue_UMB_D_B"/><tableColumn id="5" name="CM_Production"/><tableColumn id="6" name="CM_LD"/><tableColumn id="7" name="CM_UMB_D_B"/><tableColumn id="8" name="Headcount_Field"/><tableColumn id="9" name="Hours_Worked"/><tableColumn id="10" name="Warranty_Unbillable_Material"/><tableColumn id="11" name="Warranty_Unbillable_Labor_Hours"/>
</tableColumns>
<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>'''

    with zipfile.ZipFile(XLSX_NAME, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        archive.writestr("xl/styles.xml", styles_xml)
        archive.writestr("xl/worksheets/sheet1.xml", build_assumptions_sheet())
        archive.writestr("xl/worksheets/sheet2.xml", build_forecast_sheet())
        archive.writestr("xl/worksheets/sheet3.xml", build_daily_inputs_sheet())
        archive.writestr("xl/worksheets/sheet4.xml", build_scorecard_sheet())
        archive.writestr("xl/worksheets/sheet5.xml", build_capacity_sheet())
        archive.writestr("xl/worksheets/sheet6.xml", build_cashflow_sheet())
        archive.writestr("xl/worksheets/_rels/sheet3.xml.rels", sheet3_rels)
        archive.writestr("xl/tables/table1.xml", table_xml)


def build_csvs():
    daily_columns = [
        "Date",
        "Revenue_Production",
        "Revenue_LD",
        "Revenue_UMB_D_B",
        "CM_Production",
        "CM_LD",
        "CM_UMB_D_B",
        "Headcount_Field",
        "Hours_Worked",
        "Warranty_Unbillable_Material",
        "Warranty_Unbillable_Labor_Hours",
    ]
    with open(DAILY_CSV, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerow(daily_columns)

    ar_columns = [
        "Invoice_Number",
        "Customer",
        "Invoice_Date",
        "Due_Date",
        "Amount",
        "Amount_Collected",
        "Balance_Remaining",
        "Days_Outstanding",
        "Status",
        "Notes",
    ]
    with open(AR_CSV, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerow(ar_columns)


def main():
    build_xlsx()
    build_csvs()
    print(f"Generated {XLSX_NAME}, {DAILY_CSV}, and {AR_CSV}")


if __name__ == "__main__":
    main()
