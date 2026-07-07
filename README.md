# lsvtool

Local web tool for processing LSV `.DTA` / `.DTA.###` files.

## Features

- Drag and drop multiple LSV raw data files.
- Enter pH, electrode area, and sheet name for each curve.
- Convert reference potential to RHE with Ag/AgCl, Hg/HgO, or Custom reference offsets.
- Optional iR correction with Rs and compensation fraction.
- Auto, positive, or negative current sign handling for OER current direction.
- Export processed Excel and Origin-friendly tables.
- Generate LSV plots from processed C/E columns.
- Extract target-current OER overpotentials: eta100, eta200, eta300 by default.
- Fit Tafel slopes over a configurable current-density range.

## Outputs

Each export downloads a zip containing:

- `LSV_full.xlsx`: full processed workbook, one sheet per curve.
- `Origin_CE.xlsx`: Origin-friendly workbook using columns C/E.
- `Origin_CE.csv`: Origin-friendly CSV using columns C/E.
- `LSV_plot.png`: LSV curve plot.
- `Overpotential_Targets_plot.png`: overpotential-vs-current-density plot with target markers.
- `Tafel_fit.xlsx`: Tafel data and fit summary.
- `Tafel_plot.png`: Tafel plot with fit lines.

`LSV_full.xlsx` also includes:

- `Overpotential_Targets`: target j, equivalent current A, E_RHE_iRcorr, eta mV/V, pH, area, reference, iR settings, and status.
- `Overpotential_Data`: intermediate calculated j, E_RHE, E_RHE_iRcorr, and eta columns.
- `Summary`: `Eta_100_mV`, `Eta_200_mV`, `Eta_300_mV`, matching E_RHE columns, and notes.
- `Figure`: inserted overpotential plot.

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

## Target Overpotentials

Defaults:

- area: `0.25 cm2`
- pH: `13.2`
- reference: `Ag/AgCl`
- Ref to SHE: `0.197 V`
- target current densities: `100,200,300 mA/cm2`
- iR correction: `False`

Calculation:

```text
j_mA_cm2 = I_A * 1000 / area_cm2
E_RHE_V = E_ref_V + Ref_to_SHE_V + 0.0591 * pH
E_RHE_iRcorr_V = E_RHE_V - I_A * Rs_ohm * iR_fraction  (if enabled)
eta_mV = (E_RHE_iRcorr_V - 1.23) * 1000
```

The tool only interpolates inside the measured OER branch. It does not extrapolate. Targets outside the data range are reported as `out of range`.

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