# March-Scorecard-Excel-Template

codex/build-march-scorecard-excel-template-t08umh
This repository keeps source files in version control and generates the workbook artifact locally.

## Generate templates

This repository is intentionally **text-only** (no committed binary workbook).

## Generate deliverables
main
Run:

```bash
python generate_templates.py
```

## Output files
Running the command creates/refreshes:
- `March_Scorecard_Template.xlsx`
- `Daily_Inputs_Template.csv`
- `AR_Detail_Template.csv`

The two CSV templates are committed as plain text files in this repository.
