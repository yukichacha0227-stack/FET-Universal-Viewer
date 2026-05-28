import tempfile
import unittest
from pathlib import Path

from src.data_loader import merge_measurement_files, read_measurement_file, write_merged_excel


ROOT = Path(__file__).resolve().parents[1]
SAMPLE_FILES = sorted((ROOT / "sample_data").glob("*.Dat"))


class DataLoaderTest(unittest.TestCase):
    def test_sample_dat_file_is_normalized(self):
        self.assertTrue(SAMPLE_FILES, "sample_data/*.Dat is required for this test")

        df = read_measurement_file(SAMPLE_FILES[0])

        self.assertTrue({"Temp", "Mag", "Isd", "Vsd", "Vbg"}.issubset(df.columns))
        self.assertEqual(df.loc[0, "Vbg"], -15.0)
        self.assertEqual(df.loc[0, "Vsd"], -1.0)
        self.assertEqual(df.loc[0, "_SourceFile"], SAMPLE_FILES[0].name)

    def test_sample_files_can_be_merged_and_exported(self):
        self.assertGreaterEqual(len(SAMPLE_FILES), 3, "at least three sample files are required")

        result = merge_measurement_files(SAMPLE_FILES[:3])

        self.assertEqual(result.processed_count, 3)
        self.assertIn("_Sort_ID", result.dataframe.columns)
        self.assertIn("Vsd_R", result.dataframe.columns)
        self.assertIn("Vbg_R", result.dataframe.columns)
        self.assertTrue(result.dataframe["_Sort_ID"].is_monotonic_increasing)

        with tempfile.TemporaryDirectory(dir=ROOT) as tmpdir:
            output_path = Path(tmpdir) / "merged.xlsx"
            write_merged_excel(result, output_path)
            self.assertTrue(output_path.exists())


if __name__ == "__main__":
    unittest.main()
