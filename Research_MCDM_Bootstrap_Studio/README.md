# Research MCDM Bootstrap Studio

Native Python desktop application for expert-based bootstrap analysis. The interface uses Tkinter; calculations are separated from the UI so additional MCDM methods can be added without growing one monolithic screen.

## Supported methods

### DEMATEL (including fuzzy labels)

- Resamples experts with replacement, keeping each expert's full matrix intact.
- Re-averages and recalculates the normalized total influence matrix for every resample.
- Reports percentile confidence intervals for prominence (`r+c`) and relation (`r-c`), direction stability, cause/effect, and normalized prominence weight.
- Saves a results workbook and an Impact-Relation Map / bootstrap distribution chart.
- Expert sheets must be named `Ans 1`, `Ans 2`, …; the factor labels are on the configured header row and each expert sheet contains a square matrix. The diagonal is treated as zero.
- A fuzzy scale workbook has its label in column A and `L`, `M`, `U` in columns B–D. The sample file `FUZZY SCALE.xlsx` can be used as a starting point.
- The app previews each selected workbook by worksheet. For a fuzzy scale, it also shows the parsed labels and `L`, `M`, `U` values that the analysis will use.
- The defuzzification expression remains editable. Use `l`, `m`, and `u`, for example `(l + m + u)/3`, `(l + 2*m + u)/4`, `l**2`, or `m**0.5`. Use `**` for powers. The formula can be checked before running.

### TOPSIS

- Resamples complete expert rating sheets with replacement and recalculates TOPSIS closeness scores.
- Reports score confidence intervals, bootstrap mean/SD, mean rank, and the share of resamples in which each alternative ranked first.
- Expert sheets use alternatives in rows (column A) and criteria in columns (row 1 from column B). Sheets are named `Ans 1`, `Ans 2`, … .
- Scores must be non-negative. Set every criterion to `Benefit` or `Cost` in the same order as the workbook columns. Weights are optional; blank weights mean equal weights.
- A single `Data` sheet is accepted for data inspection but cannot be bootstrapped; at least two expert sheets are required.

### AHP

- Resamples experts and combines each sampled panel with an elementwise geometric mean of reciprocal pairwise matrices.
- Reports group and bootstrap priority weights, percentile intervals, rank stability, each expert's consistency ratio (CR), and group/bootstrap CR summaries.
- Uses Saaty random-index values through 15 criteria. The app flags the common `CR ≤ 0.10` guideline; interpret it as a diagnostic, not an automatic guarantee of valid judgments.
- Input sheets use the DEMATEL matrix layout (`Ans 1`, `Ans 2`, …), but the pairwise matrix must be numeric, positive, reciprocal, and have a diagonal of 1. Fuzzy labels are not accepted for AHP in this version.

## Run on Windows

1. Install Python 3.10 or newer and keep the Tcl/Tk option enabled in the Python installer.
2. Open this folder in File Explorer.
3. Run `run_app.bat`.

If dependencies are missing, open Command Prompt in this folder and run:

```bat
python -m pip install -r requirements.txt
python run_app.py
```

The app writes one `.xlsx` results file and one `.png` chart to the selected output folder. Existing files are not overwritten because result names include a timestamp.

## Build a Windows executable

Run `build_app.bat` on Windows after installing the requirements. PyInstaller is installed by the script and writes the app under `dist/ResearchMCDMBootstrap/`.

## Statistical interpretation

The DEMATEL column `Direction stability` is the fraction of bootstrap resamples retaining the original sign. It is descriptive and is not a p-value. A percentile interval that excludes zero indicates a stable direction under this resampling procedure; it is not a causal claim or a substitute for a prespecified inferential model. Bootstrap uncertainty reflects variation across the supplied experts and does not correct expert-selection bias.

TOPSIS intervals reflect variation in the resampled expert panel. They do not capture uncertainty from criterion selection, weight elicitation, or the choice of normalization method. Check sensitivity to weights and Benefit/Cost assumptions as part of the analysis.

## Suggested next upgrades

1. Add AHP bootstrap with a consistency ratio (CR) report and an option to exclude or repair inconsistent expert matrices.
2. Add Delphi consensus stability, such as agreement/dispersion by round and threshold sensitivity.
3. Add fuzzy DEMATEL options for alternative defuzzification, Z-number confidence handling, and a side-by-side crisp/fuzzy sensitivity report.
4. Add workbook mapping and validation for labels, missing cells, duplicate names, and matrix diagonals.
5. Add comparison runs for seeds, sample counts, CI methods, weights, thresholds, and Benefit/Cost settings.
6. Improve visual analysis with an editable IRM threshold, quadrant labels, overlap-aware factor labels, SVG/PDF output, and reproducible figure settings.
7. Add a project file that stores workbook paths, method settings, run history, and output locations so an analysis can be reopened exactly.

## Files

- `run_app.py`, `app.py`: native desktop entry point and interface.
- `bootstrap_dematel.py`: fuzzy DEMATEL calculations and export.
- `bootstrap_topsis.py`: TOPSIS calculations and export.
- `requirements.txt`: calculation dependencies. Tkinter is included with standard Windows Python installations.
- `templates/`: example workbooks for DEMATEL, AHP, TOPSIS, and a fuzzy scale.
