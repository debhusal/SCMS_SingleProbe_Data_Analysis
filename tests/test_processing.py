import csv
import io
import math
import random
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from openpyxl import Workbook
from processing import (Settings, SpectrumError, Exclusions, background_ranges,
                        legacy_boundary, peaks, process_workbook, sheet_names)


def make_workbook(sheets):
    book = Workbook()
    book.remove(book.active)
    for name, rows in sheets:
        sheet = book.create_sheet(name)
        for row in rows:
            sheet.append(row)
    stream = io.BytesIO()
    book.save(stream)
    book.close()
    stream.seek(0)
    return stream


class LegacyTests(unittest.TestCase):
    def test_percent_format_exact(self):
        for mass in [100, 493.2511, 0.001, 99999.99995, 123.456789]:
            self.assertEqual(legacy_boundary(mass), (float("%8.4f" % (mass * 0.999995)),
                                                     float("%8.4f" % (mass * 1.000005))))

    def test_strict_endpoints_and_touching_ranges(self):
        index = Exclusions([(1, 2), (2, 3), (4, 4), (5, 7), (6, 8)])
        for value in [1, 2, 3, 4, 5, 8]:
            self.assertFalse(index.contains(value))
        for value in [1.5, 2.5, 6, 7]:
            self.assertTrue(index.contains(value))
        self.assertTrue(index.contains(math.nextafter(2, 1)))
        self.assertTrue(index.contains(math.nextafter(2, 3)))

    def test_index_matches_literal_boolean(self):
        rng = random.Random(42)
        ranges = [(rng.randint(0, 100) / 10, rng.randint(101, 200) / 10) for _ in range(100)]
        ranges += [(21, 22), (22, 23), (24, 24)]
        index = Exclusions(ranges)
        values = [rng.uniform(0, 25) for _ in range(1000)] + [v for pair in ranges for v in pair]
        for value in values:
            self.assertEqual(index.contains(value), any(lo < value < hi for lo, hi in ranges))

    def test_text_formats_and_cutoff(self):
        for delimiter in [",", "\t", ";", " "]:
            for encoding in ["utf-8-sig", "utf-16", "cp1252"]:
                text = "Thermo export\n" + delimiter.join(["m.z", "Relative", "Intensity"]) + "\n"
                text += delimiter.join(["100", "99", "3000"]) + "\n"
                text += delimiter.join(["200", "100", "2999"]) + "\n"
                stream = io.BytesIO(text.encode(encoding))
                ranges, before = background_ranges(stream, "background.txt", Settings())
                self.assertEqual(before, 2)
                self.assertEqual(ranges, [(99.9995, 100.0005)])
                self.assertFalse(stream.closed)

    def test_end_to_end_denominator_headers_padding_order(self):
        sample = make_workbook([
            ("Cell Z", [["metadata"], [], ["m/z", "Relative", "Intensity"],
                        [100, 1, 9999], [99.9995, 1, 3000], [100.0005, 1, 3000],
                        [200, 1, 2000], [300, 1, 6000], [300, 1, 6000]]),
            ("Cell A", [["Intensity", "m.z"], [6000, 500], [2000, 600]]),
            ("Cell C", [["M.Z", "Intensity"], [700, 3000]])])
        self.assertEqual(sheet_names(sample), ["Cell Z", "Cell A", "Cell C"])
        with tempfile.TemporaryDirectory() as temp:
            summary = process_workbook(sample, [legacy_boundary(100)], Settings(), temp)
            first = summary[0]
            self.assertEqual([first[k] for k in ["peaks_before", "removed_background", "removed_intensity", "remaining"]], [6, 1, 1, 4])
            self.assertEqual(first["normalization_denominator"], 20000)
            with (Path(temp) / "POS_legacy.csv").open(newline="") as f:
                rows = list(csv.reader(f))
            self.assertEqual(rows[0], ["1", "1", "2", "2", "3", "3"])
            self.assertEqual(rows[1], ["mass", "abund"] * 3)
            self.assertEqual(rows[2], ["99.9995", "15000", "500", "75000", "700", "100000"])
            self.assertEqual(rows[3], ["100.0005", "15000", "NA", "NA", "NA", "NA"])
            self.assertEqual(rows[4], rows[5])  # duplicate peaks preserved
            with (Path(temp) / "POS_tidy.csv").open(newline="") as f:
                tidy = list(csv.DictReader(f))
            self.assertEqual([r["sheet"] for r in tidy], ["Cell Z"] * 4 + ["Cell A", "Cell C"])

    def test_validation(self):
        cases = [([], "empty"), ([["m.z", "Relative"]], "missing"),
                 ([["m.z", "Intensity"]], "no spectrum"),
                 ([["m.z", "Intensity"], [100, "bad"]], "row 2"),
                 ([["m.z", "Intensity"], [100, float("nan")]], "row 2")]
        for rows, message in cases:
            with self.subTest(message=message), self.assertRaisesRegex(SpectrumError, message):
                list(peaks(rows, "test"))

    def test_zero_and_no_retained_and_no_partial_export(self):
        for value, message in [(0, "zero"), (1000, "no peaks remain")]:
            sample = make_workbook([("valid", [["m.z", "Intensity"], [200, 5000]]),
                                    ("bad", [["m.z", "Intensity"], [200, value]])])
            with tempfile.TemporaryDirectory() as temp:
                with self.assertRaisesRegex(SpectrumError, message):
                    process_workbook(sample, [], Settings(), temp)
                self.assertEqual(list(Path(temp).iterdir()), [])

    def test_neg_labels_and_background_selection(self):
        background = make_workbook([("one", [["m.z", "Intensity"], [100, 4000]]),
                                    ("two", [["m.z", "Intensity"], [200, 4000]])])
        with self.assertRaisesRegex(SpectrumError, "Choose one"):
            background_ranges(background, "bg.xlsx", Settings())
        ranges, _ = background_ranges(background, "bg.xlsx", Settings(), "two")
        self.assertEqual(ranges, [legacy_boundary(200)])
        sample = make_workbook([("cell", [["m.z", "Intensity"], [300, 4000]])])
        with tempfile.TemporaryDirectory() as temp:
            process_workbook(sample, ranges, Settings(ion_mode="NEG"), temp)
            self.assertTrue(all(p.name.startswith("NEG_") for p in Path(temp).iterdir()))

    def test_existing_outputs_protected(self):
        with tempfile.TemporaryDirectory() as temp:
            marker = Path(temp) / "POS_legacy.csv"
            marker.write_text("existing")
            with self.assertRaisesRegex(SpectrumError, "empty"):
                process_workbook(io.BytesIO(), [], Settings(), temp)
            self.assertEqual(marker.read_text(), "existing")


if __name__ == "__main__":
    unittest.main()
