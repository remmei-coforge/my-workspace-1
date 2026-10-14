
# Python Excel Field Mapping Agent

A configurable Streamlit application that maps source Excel workbook values into a copy of the BFS Resource Tracker template.

## Features

- Multiple `.xlsx` and `.xlsm` uploads.
- Configurable worksheet names, header rows, source/target keys, and column mappings in `mapping-setting.json`.
- Key normalization by trimming surrounding whitespace.
- Duplicate and blank key handling.
- Fill-blank-only or overwrite policy (mutually exclusive).
- Avoids formula overwrites unless explicitly enabled.
- Reconciliation report for matched, unmatched, blank-key, and duplicate records.
- Creates a unique timestamped output workbook and verifies it can be reopened.
- Uses `openpyxl` to preserve workbook structures and formatting wherever possible.

## Requirements

- Python 3.11+
- Windows, macOS, or Linux

## Setup (Windows / PowerShell)

```powershell
cd excel-field-mapping-agent
py -3.13 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
streamlit run app.py
```

If PowerShell blocks activation, use Command Prompt:

```bat
.venv\Scripts\activate.bat
pip install -r requirements.txt
streamlit run app.py
```

## Try the sample data

Upload `sample-data/Dummy-Source-Data.xlsx` in the app. The source worksheets and headers should correspond to the settings in `mapping-setting.json`.

## Configure mappings

`mapping-setting.json` controls:
- `source_sheet` / `target_sheet`
- `source_header_row` / `target_header_row`
- `source_key` / `target_key`
- `mapping`: source column name to target column name
- `fill_blank_cells_only` and `overwrite_existing_values`
- `fail_on_duplicate_keys`
- `allow_formula_overwrite`

Do not enable both fill-blank-only and overwrite-existing-values.

## Important notes

1. The included configuration uses `Opportunity` as the Open Demand Tracker key and `Emp Code` as the Resource Release List key, because `Emp Code` is present in the provided target template.
2. Source records with no matching target key are skipped by default and reported. This starter version does not append new target records.
3. `.xlsm` sources can be read. The provided target template is `.xlsx`; output is therefore `.xlsx`.
4. `openpyxl` does not preserve every advanced Excel feature (for example, some external data connections, slicers, or embedded objects). Test with a copy of production workbooks before operational use.
5. This is a starter project. Review the JSON mapping against real source headers before using production data.
