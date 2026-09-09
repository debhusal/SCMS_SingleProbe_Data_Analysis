# Validation performed

Date: 2026-09-08.

## Reference comparison

The supplied R script's Boolean expression and processing block were copied into a scratch test harness. Only the input/output setup was replaced: the harness used three generated numeric spectra, base R CSV input instead of readxl, and a scratch output file. The production app uses the uploaded background, rather than the original experiment's hardcoded ranges.

R 4.5.0 comparison results:

* 1,185 original background intervals were tested.
* Three spectra included interval endpoints, midpoints, nearby floating-point values, and sample intensities above, at, and below 3000. R and Python received the same Excel-round-tripped numeric values.
* Legacy output: 2,312 rows including both header rows.
* Numeric differences: zero in all exported fields.
* Text-field differences: zero, including NA padding.
* Full files matched after normalizing operating-system line endings.

The original `format.py` was executed unchanged on five test masses, including small masses, large masses, and fractional values. The generated four-decimal boundary text matched exactly.

This is a synthetic regression comparison, not validation against an actual laboratory workbook or every possible R floating-point/string-formatting case. Missing/invalid numeric rows intentionally produce explicit validation errors, as documented in README.md.

## Other checks

Nine automated processing tests passed: four-decimal boundaries; strict endpoints and touching intervals; indexed exclusion equivalence to a literal Boolean scan; supported text delimiters/encodings; denominator, ordering, duplicate retention, two headers, and padding; malformed/empty inputs; zero/all-filtered cases and partial-export cleanup; background sheet selection and NEG naming; existing-output protection.

Streamlit AppTest checks passed using injected in-memory uploads: initial rendering, processing disabled without inputs, successful workbook processing, summary rendering, POS and NEG exports, download preparation, and visible validation errors. Test dependencies: Streamlit 1.63.0, openpyxl 3.1.5, Python 3.14.6. Core tests also passed under Python 3.12.

## Original files unchanged

SHA-256 values were identical before and after implementation:

* `format.py`: `C22FA17B44922DDC80D592C5897D8E4CA8BBD82DE3603C079A7A529E5857AD93`
* `code for data analysis.R`: `1E12A4D188001B81882B1C36C48E0B72AF49E4B0DC0099259464F76319ED4427`
