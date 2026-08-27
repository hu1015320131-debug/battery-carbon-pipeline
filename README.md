# battery-carbon-pipeline｜电芯碳数据标准化与核算工具

Local-first Excel import, cleaning, matching and carbon-calculation demo.

This public repository uses **synthetic data only**. It is not a production
inventory system and does not ship enterprise ledgers or real emission factors.

## What it does

- Inspect an `.xlsx` file in an isolated run directory
- Detect headers and map fields from aliases
- Keep non-cell rows out of cell calculation
- Route activity (PCS × unit weight, or reported mass) from capability, not from year
- Calculate `Activity × EF` with Python `Decimal`
- Show a local Streamlit UI and download results

Unknown files do **not** inherit a default factor. Example factor `1.250000 kgCO2e/kg`
is a synthetic fixture (`production_eligible = false`).

## Requirements

- Python 3.11+
- Windows, macOS or Linux

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e ".[dev]"
$env:PYTHONPATH = "src"
.\.venv\Scripts\python.exe -m carbon_excel_pipeline check
.\.venv\Scripts\python.exe -m pytest
```

On Unix:

```bash
python -m venv .venv
.venv/bin/python -m pip install -e ".[dev]"
PYTHONPATH=src .venv/bin/python -m carbon_excel_pipeline check
PYTHONPATH=src .venv/bin/python -m pytest
```

## Run the example

```powershell
$env:PYTHONPATH = "src"
.\.venv\Scripts\python.exe -m carbon_excel_pipeline wp6-8-1-e2e `
  --input examples/public/synthetic_cells.xlsx `
  --run-root $env:TEMP\carbon-excel-runs
```

## Streamlit UI

```powershell
$env:PYTHONPATH = "src"
Copy-Item config/local_paths.example.json config/private/local_paths.json
.\.venv\Scripts\python.exe -m streamlit run app/streamlit_app.py --server.address 127.0.0.1 --server.port 8501
```

Point `run_root` in `config/private/local_paths.json` to a folder outside Git.
That file is gitignored.

## Safety

See [NOTICE](NOTICE). Do not publish run outputs that contain real company files.
Do not treat UI totals as year-on-year corporate reductions.

## License

MIT. See [LICENSE](LICENSE).
