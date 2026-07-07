# lsvtool

Local web tool for processing LSV `.DTA` / `.DTA.###` files.

## Features

- Drag and drop multiple LSV raw data files.
- Enter pH, electrode area, and sheet name for each curve.
- Export processed Excel and Origin-friendly tables.
- Generate LSV plots from the processed C/E columns.
- Fit Tafel slopes over a configurable current-density range.

## Outputs

Each export downloads a zip containing:

- `LSV_full.xlsx`: full A-E processed workbook, one sheet per curve.
- `Origin_CE.xlsx`: Origin-friendly workbook using columns C/E.
- `Origin_CE.csv`: Origin-friendly CSV using columns C/E.
- `LSV_plot.png`: LSV curve plot.
- `Tafel_fit.xlsx`: Tafel data and fit summary.
- `Tafel_plot.png`: Tafel plot with fit lines.

## Install

Install Python 3.10 or newer, then install dependencies:

```powershell
pip install -r requirements.txt
```

## Run

On Windows, double-click:

```text
start_lsv_tool.bat
```

Or run manually:

```powershell
python scripts\lsv_tool_app.py
```

Then open:

```text
http://127.0.0.1:8765
```

## Tafel Fit

The Tafel fit uses:

```text
eta = (E_RHE - 1.23) * 1000
eta = slope * log10(j) + intercept
```

- `E_RHE` comes from column C of the processed LSV table.
- `j` comes from column E, in `mA/cm2`.
- Slope is reported in `mV/dec`.
- Default fitting range is `10` to `100 mA/cm2`.

## Notes

- Files are processed locally. The app does not upload data to any external server.
- The no-Python portable runtime package is intentionally not included in this repository.
- Generated outputs are ignored by Git.
