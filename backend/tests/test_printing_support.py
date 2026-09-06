import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import app


class PrintingSupportTests(unittest.TestCase):
    def test_get_win32_support_reports_current_python(self):
        support = app.get_win32_support()

        self.assertIn("available", support)
        self.assertIn("python_executable", support)
        self.assertEqual(support["python_executable"], sys.executable)

        if support["available"]:
            self.assertTrue(callable(support["Dispatch"]))
            self.assertIsNotNone(support["SW_SHOWNORMAL"])

    def test_imprimir_via_shell32_returns_value_when_available(self):
        path = Path(__file__)
        result = app.imprimir_via_shell32(path)
        self.assertIsNone(result) if sys.platform != "win32" else self.assertIsNotNone(result)

    def test_get_printers_returns_list(self):
        printers = app.get_printers()
        self.assertIsInstance(printers, list)
        self.assertTrue(all(isinstance(item, str) and item for item in printers))

    def test_temp_dir_points_to_downloads_temp(self):
        self.assertEqual(app.get_temp_dir(), Path.home() / "Downloads" / "temp")

    def test_set_default_printer_is_noop_on_non_windows(self):
        self.assertFalse(app.set_default_printer("Dummy Printer"))


if __name__ == "__main__":
    unittest.main()
