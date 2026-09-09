"""Run with: python -m streamlit run app.py"""
import json
import shutil
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
from uuid import uuid4

import streamlit as st

from processing import Settings, SpectrumError, background_ranges, process_workbook, sheet_names


st.set_page_config(page_title="Single-Probe preprocessing", page_icon="🧪", layout="centered")
st.title("Single-Probe preprocessing")
st.write("Remove background peaks and normalize single-cell spectra using your legacy workflow.")
st.info("Legacy normalization: total intensity is calculated after background removal, BEFORE the sample intensity cutoff, exactly as in your R script.")
mode = st.selectbox("Ion mode for this run", ["POS", "NEG"], help="Labels filenames only. Upload background and cells from the same ion mode.")
background = st.file_uploader("1. Background spectrum", type=["xlsx", "xlsm", "csv", "txt", "tsv"])
bg_sheet = None
bg_ok = True
if background is not None and Path(background.name).suffix.lower() in (".xlsx", ".xlsm"):
    try:
        bg_sheet = st.selectbox("Background sheet", sheet_names(background))
    except SpectrumError as exc:
        st.error(str(exc))
        bg_ok = False
samples = st.file_uploader("2. Single-cell workbook (one cell per sheet)", type=["xlsx", "xlsm"])
st.caption("All workbook sheets are processed in workbook order, including hidden sheets. Use a separate workbook/run for each ion mode. Older .xls files: Save As .xlsx in Excel first.")
with st.expander("Settings — defaults match the legacy workflow"):
    bg_cutoff = st.number_input("Background intensity cutoff", min_value=0.0, value=3000.0, step=100.0)
    sample_cutoff = st.number_input("Sample intensity cutoff", min_value=0.0, value=3000.0, step=100.0)
    ppm = st.number_input("Background tolerance (± ppm)", min_value=0.001, max_value=999999.0, value=5.0, step=1.0, format="%.3f")
    factor = st.number_input("Normalization factor", min_value=1.0, value=100000.0, step=1000.0)
    st.caption("Boundaries use %8.4f formatting. Only peaks strictly inside a range are removed. Peaks exactly at the intensity cutoff are retained.")
settings = Settings(bg_cutoff, sample_cutoff, ppm, factor, mode)

if st.button("Process workbook", type="primary", disabled=background is None or samples is None or not bg_ok):
    st.session_state.pop("result", None)
    status = st.empty()
    bar = st.progress(0.0)
    run_dir = None
    try:
        ranges, bg_count = background_ranges(background, background.name, settings, bg_sheet)
        if not ranges:
            st.warning("No background peaks meet the cutoff. This run will apply no background exclusions.")
        result_root = Path(__file__).resolve().parent / "results"
        result_root.mkdir(exist_ok=True)
        run_dir = result_root / f"{mode}_{datetime.now():%Y%m%d_%H%M%S}_{uuid4().hex[:6]}"
        # Only publish the run directory once every sheet passes validation.
        with TemporaryDirectory(prefix="processing_", dir=result_root) as scratch:
            def progress(done, total, name):
                bar.progress(done / total)
                status.write(f"{done}/{total} cells complete — {name}")

            summary = process_workbook(samples, ranges, settings, scratch, progress)
            manifest = Path(scratch) / f"{mode}_settings.json"
            details = json.loads(manifest.read_text(encoding="utf-8"))
            details.update({"background_file": background.name, "background_sheet": bg_sheet,
                            "sample_file": samples.name, "background_peaks_before": bg_count,
                            "background_peaks_retained": len(ranges)})
            manifest.write_text(json.dumps(details, indent=2), encoding="utf-8")
            run_dir.mkdir()
            for path in Path(scratch).iterdir():
                shutil.copyfile(path, run_dir / path.name)
        st.session_state.result = {"folder": str(run_dir), "summary": summary, "mode": mode,
                                   "background_count": len(ranges), "sample_file": samples.name}
        status.empty()
    except (SpectrumError, OSError, ValueError, OverflowError) as exc:
        st.error(f"Run not exported: {exc}")
        bar.empty()
        status.empty()
        if run_dir and run_dir.exists():
            # Only files in this newly created, app-owned run folder.
            for path in run_dir.iterdir():
                path.unlink()
            run_dir.rmdir()

if "result" in st.session_state:
    result = st.session_state.result
    st.success(f"Completed {result['mode']} run: {result['sample_file']} — {len(result['summary'])} cells.")
    st.caption("Results below belong to this completed run. Changing uploads or settings requires clicking Process workbook again.")
    st.write(f"Background peaks retained: {result['background_count']}")
    st.dataframe(result["summary"], hide_index=True)
    st.write("Results saved on this computer:")
    st.code(result["folder"], language=None)
    folder = Path(result["folder"])
    st.caption("Legacy CSV contains two header rows, alternating mass/abund columns, and NA padding for shorter cells. The summary maps cell numbers to sheet names.")
    # Avoid eagerly loading all exports into Streamlit's download memory.
    choices = sorted(p.name for p in folder.glob("*") if p.is_file())
    chosen = st.selectbox("File to download", choices)
    if st.checkbox("Prepare selected download", help="Optional: files are already saved locally. Large downloads use additional RAM."):
        path = folder / chosen
        if path.stat().st_size > 25 * 1024 * 1024:
            st.warning("This file exceeds 25 MB. Open it from the results folder shown above to avoid an additional RAM copy.")
        else:
            with path.open("rb") as data:
                st.download_button("Download selected file", data, file_name=path.name, mime="application/json" if path.suffix == ".json" else "text/csv")
