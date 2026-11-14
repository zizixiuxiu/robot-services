import base64
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

SRC_DIR = Path(__file__).resolve().parents[1] / "src"
sys.path.insert(0, str(SRC_DIR))

from workshop_order_http import (  # noqa: E402
    _batch_zip_outputs,
    _flatten_output_items,
    _unique_arcname,
)


class BatchZipTests(unittest.TestCase):
    def test_unique_arcname_suffix(self) -> None:
        used: set[str] = set()
        self.assertEqual(_unique_arcname("a.xlsx", used), "a.xlsx")
        self.assertEqual(_unique_arcname("a.xlsx", used), "a(2).xlsx")
        self.assertEqual(_unique_arcname("a.xlsx", used), "a(3).xlsx")

    def test_single_output_not_zipped(self) -> None:
        response = {
            "output_files": {
                "a.xls": [{"filename": "a下车间.xlsx", "file_content": base64.b64encode(b"one").decode()}],
            }
        }
        _batch_zip_outputs(response, [], [])
        self.assertEqual(len(_flatten_output_items(response["output_files"])), 1)
        self.assertNotIn("zip_original_count", response)

    def test_multi_output_zipped(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            p1 = Path(tmp) / "甲下车间.xlsx"
            p2 = Path(tmp) / "乙下车间.xlsx"
            p1.write_bytes(b"AAA")
            p2.write_bytes(b"BBB")
            response = {
                "output_files": {
                    "甲.xls": [{"filename": p1.name, "path": str(p1), "file_content": "ignored"}],
                    "乙.xls": [{"filename": p2.name, "path": str(p2), "file_content": "ignored"}],
                }
            }
            _batch_zip_outputs(response, [], [])
            items = _flatten_output_items(response["output_files"])
            self.assertEqual(len(items), 1)
            self.assertEqual(response.get("zip_original_count"), 2)
            zip_name = items[0]["filename"]
            self.assertTrue(zip_name.startswith("下车间批量_"))
            self.assertTrue(zip_name.endswith("_2个.zip"))
            raw = base64.b64decode(items[0]["file_content"])
            with zipfile.ZipFile(Path(tmp) / "out.zip", "w") as _:
                pass
            out = Path(tmp) / "check.zip"
            out.write_bytes(raw)
            with zipfile.ZipFile(out) as zf:
                names = sorted(zf.namelist())
                self.assertEqual(names, ["乙下车间.xlsx", "甲下车间.xlsx"])
                self.assertEqual(zf.read("甲下车间.xlsx"), b"AAA")
                self.assertEqual(zf.read("乙下车间.xlsx"), b"BBB")


if __name__ == "__main__":
    unittest.main()
