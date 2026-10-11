import csv
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree


ROOT = Path(__file__).resolve().parents[1]


class AssetTests(unittest.TestCase):
    def test_dashboard_is_a_valid_excel_package(self):
        workbook = ROOT / "Dashboard.xlsx"
        self.assertTrue(zipfile.is_zipfile(workbook))

        with zipfile.ZipFile(workbook) as archive:
            names = set(archive.namelist())
            self.assertIn("[Content_Types].xml", names)
            self.assertIn("xl/workbook.xml", names)
            root = ElementTree.fromstring(archive.read("xl/workbook.xml"))
            namespace = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            sheets = root.findall("x:sheets/x:sheet", namespace)
            self.assertGreaterEqual(len(sheets), 1)

    def test_source_csv_has_rows_and_required_columns(self):
        with (ROOT / "data.csv").open(newline="", encoding="utf-8-sig") as handle:
            reader = csv.DictReader(handle)
            self.assertTrue({"Sales", "Profit", "Region"}.issubset(reader.fieldnames or []))
            self.assertGreater(sum(1 for _ in reader), 100)


if __name__ == "__main__":
    unittest.main()
