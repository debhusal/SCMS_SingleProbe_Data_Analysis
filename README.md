# Single-Probe preprocessing

A local Streamlit app for Thermo Xcalibur background exports and one-cell-per-sheet Excel workbooks. No database, Docker, R installation, cloud service, or new peak-processing algorithm is required. Only Streamlit and openpyxl are direct dependencies; Streamlit installs its own dependencies (including pandas).

## Windows: download and start

1. Install **64-bit Python 3.12 or newer** from [python.org](https://www.python.org/downloads/windows/) if needed. Include the Python launcher and select **Add Python to PATH**.
2. Open [this repository](https://github.com/debhusal/SCMS_SingleProbe_Data_Analysis), click the green **Code** button, then **Download ZIP**. No Git installation is needed.
3. Right-click the downloaded ZIP and select **Extract All**. Open the extracted folder containing `app.py`, `install.bat`, and `launch.bat`. Do not run the files from inside the ZIP.
4. Double-click **install.bat** once. It creates a local `.venv` and installs the requirements. Internet access is needed for this step. Wait for **Installation complete**, then close that command window.
5. Double-click **launch.bat**. The application opens in your browser. If Streamlit asks for an email, leave it blank and press **Enter**; an account is not required.
6. If the browser does not open automatically, visit [http://127.0.0.1:8501](http://127.0.0.1:8501).

Keep the command window open while using the application. Press **Ctrl+C** in that window to stop it. On subsequent uses, run **launch.bat**; you do not need to run the installer again unless requirements change. After installation, processing works offline. **R is not required.**

If the installer cannot find Python, install Python with its launcher, then retry. If installation fails, check the internet connection and the error shown in the command window. If you move the application folder and launching fails, delete only its `.venv` folder and rerun `install.bat`.

Optional command-line installation from the extracted repository folder:

```bat
py -m venv .venv
.venv\Scripts\python.exe -m pip install --no-cache-dir -r requirements.txt
.venv\Scripts\python.exe -m streamlit run app.py
```

## Repository contents

- `app.py`: Streamlit GUI.
- `processing.py`: streaming parsers, background filtering, normalization, and exports.
- `requirements.txt`, `install.bat`, `launch.bat`: Windows setup and launch.
- `.streamlit/config.toml`: local-only server, upload limit, and usage statistics disabled.
- `tests/`: synthetic regression tests; no personal datasets.
- `legacy_reference/`: original formatting script and a sanitized copy of the legacy R workflow, for reference only. The GUI does not import or execute them.
- `VALIDATION.md`: validation scope and results.

Datasets, run results, temporary files, credentials, and virtual environments are excluded. `.gitignore` also excludes common data formats and generated output folders.

## Using the app

1. Choose POS or NEG. This only labels the run: it cannot infer or validate the ion mode in a spectrum. Supply matching-mode background and sample data; do not put POS and NEG cells in one workbook.
2. Upload a background `.xlsx`, `.xlsm`, `.csv`, `.txt`, or `.tsv`. Choose one sheet if the background is Excel. Excel `.xls` is not supported: open it in Excel and Save As `.xlsx` first.
3. Upload the sample `.xlsx` or `.xlsm` workbook. Every worksheet is processed in workbook order, including hidden sheets. No sheet count is needed.
4. Keep the default settings and click **Process workbook**.
5. Review counts and open the saved results folder. Downloads are optional. Each run has a new timestamped POS/NEG directory inside `results`, beside `app.py`.

Header detection accepts case/punctuation/spacing variants such as `m.z`, `m/z`, `MZ`, and `Intensity`. `Relative` is ignored. Metadata before the header, blank rows, and repeated spectrum headers are ignored. Text exports support comma, tab, semicolon, or whitespace separators and UTF-8 BOM, UTF-16 BOM, or Windows-1252 text. Use decimal points and ungrouped numeric values (scientific notation is accepted).

## Exact reference rules and an important denominator difference

The implementation was checked against the supplied `format.py` and `code for data analysis.R`. Copies are in `legacy_reference`; the R copy has personal paths, experiment filenames, and dataset-specific background ranges removed. The original local files were not edited. The app implements their processing rules for valid finite numeric spectra, with these defaults:

* Background intensities **>= 3000** produce ranges. This background selection step is specified by the user; neither supplied script performs that selection.
* At ±5 ppm, calculate `mass * 0.999995` and `mass * 1.000005`, then convert each using **`float("%8.4f" % value)`**. Do not round the sample mass.
* Remove a peak only when **`lower < mass < upper`** for at least one range. Exact endpoints survive unless inside another range. Duplicates and original peak order are preserved; no alignment, binning, sorting of sample peaks, or merging of peaks is performed.
* Compute the denominator from **all intensities remaining after background removal**, including intensities below 3000. This is the actual order in the supplied R script, and is preserved by the GUI.
* Retain sample peaks with intensity **>= 3000**, and calculate `(intensity / denominator) * 100000`.

Example: after background removal, intensities of 6000 and 2000 give a denominator of 8000. Only the 6000 peak is exported, with normalized abundance **75000**, not 100000. Retained normalized abundances may therefore sum to less than 100000. The summary reports this denominator explicitly.

The R script has a hardcoded background Boolean expression for its original dataset. This app generates the equivalent expression's ranges from the uploaded background; it does not silently apply the original dataset's background to new experiments.

## Exports

* **POS_legacy.csv / NEG_legacy.csv:** exact requested two-header-row structure, cell IDs starting at 1, alternating `mass,abund` columns, and no extra row index or column names. Unequal cell lengths are padded with `NA`. As with R's `rbind` with text headers and default `write.table`, character fields are quoted and padding `NA` is unquoted. The logical first rows for three cells are:

```csv
"1","1","2","2","3","3"
"mass","abund","mass","abund","mass","abund"
```

* **tidy.csv:** ion mode, cell ID, sheet name, original mass, raw intensity, and normalized intensity; one retained peak per row.
* **summary.csv:** cell ID, sheet, counts before filtering, removed as background, removed by cutoff, remaining, and normalization denominator. Background removals are counted first, so peaks removed as background are not counted again under the intensity cutoff.
* **background_ranges.csv:** every retained background range, formatted to four decimals for inspection.
* **settings.json:** settings, input filenames, sheet order, and denominator/boundary rules for that run.

Legacy numeric fields use 15 significant digits to follow R's numeric-to-character conversion; tidy values retain Python float precision. Cross-language floating-point sums and textual exponent choices may differ in the last digits. Byte-for-byte R export equivalence requires checking against the user's R version and a real reference output; the regression tests verify filtering and layout, not every possible R serialization case.

A direct synthetic comparison against the supplied R logic under R 4.5.0 passed: all 1,185 hardcoded ranges, three cells, and 2,312 CSV rows matched exactly after normalizing line endings. The original `format.py` was also executed unchanged and matched the generated boundary strings. See `VALIDATION.md` for scope and results.

## Validation and memory

Empty sheets, missing/ambiguous headers, invalid numeric rows, zero/nonfinite denominators, or cells with no retained peaks stop the run with a sheet-specific message. No partial run is published. Unlike R's permissive NA handling and its failing `1:0` loop on empty output, this app rejects missing/nonfinite numeric values explicitly instead of producing ambiguous data. Re-save formula cells with calculated values in Excel if required. Remove non-spectrum sheets or footer text before processing.

Excel is read in openpyxl read-only mode and cells are processed sequentially. Peaks are spooled to temporary disk files rather than collected in one giant DataFrame. The legacy table is streamed a row at a time. Overlapping exclusion intervals are indexed without changing strict endpoint semantics. There is no dense peaks-by-background matrix.

Streamlit still holds uploaded files in RAM, and openpyxl can hold shared strings and workbook metadata. The default upload limit is 100 MB per file, not a guarantee of fitting any workbook in RAM. Avoid oversized workbooks on a limited-memory laptop. Large exports are available on disk without loading them into download memory. Optional downloads are limited to 25 MB. See `.streamlit/config.toml` for the upload limit. Finished runs persist; delete unwanted run folders in File Explorer to reclaim space. Temporary processing files are removed after successful runs or handled errors; after a forced shutdown, any `results/processing_*` folder can be deleted while the app is stopped.

The server binds to 127.0.0.1 and Streamlit usage statistics are disabled. Processing works offline after installation. Inputs are not modified.

## Tests

```bat
.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Implementation references: [openpyxl read-only mode](https://openpyxl.readthedocs.io/en/stable/optimized.html) and [Streamlit upload limits](https://docs.streamlit.io/develop/api-reference/widgets/st.file_uploader).
