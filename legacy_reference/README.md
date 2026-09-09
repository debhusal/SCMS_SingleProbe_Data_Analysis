# Legacy workflow reference

These files document the workflow reproduced by the Python GUI. **Do not run them to launch the application. R is not needed by the GUI.**

- `format.py` is an unchanged copy of the original script. It uses NumPy, multiplies background masses by 0.999995 and 1.000005, and writes strict R comparisons using `%8.4f` boundaries.
- `code for data analysis.R` is a sanitized reference copy. The personal working-directory path and experiment filenames were replaced, and the 1,185 experiment-specific background intervals were replaced with one clearly marked synthetic example. The filtering, normalization order, and two-row output-header logic are retained. It is not a ready-to-run analysis for new data: its workbook, sheet count, and background expression are placeholders.

The original local source files were not changed. The GUI generates ranges from each uploaded background, detects all sheets, and implements the workflow entirely in Python. The optional legacy scripts have their own R/NumPy requirements; these are not additional GUI requirements.
