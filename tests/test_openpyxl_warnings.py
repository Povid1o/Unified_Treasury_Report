"""Предупреждение openpyxl о x14-расширениях не должно доходить до консоли.

Выгрузки банковских систем несут выпадающие списки в блоке extLst/x14, и
openpyxl на каждом чтении такого файла печатал «Data Validation extension is
not supported and will be removed». Фильтр стоит в common/__init__.py.
"""
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

import openpyxl

BASE_DIR = Path(__file__).resolve().parents[1]

X14_EXT = (
    '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" '
    'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
    '<x14:dataValidations count="0" '
    'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"/></ext></extLst>'
)


class OpenpyxlExtensionWarningTests(unittest.TestCase):
    def test_reading_file_with_x14_validation_is_silent(self):
        with tempfile.TemporaryDirectory() as tmp:
            plain, path = Path(tmp) / "plain.xlsx", Path(tmp) / "x14.xlsx"
            wb = openpyxl.Workbook()
            wb.active["A1"] = 1
            wb.save(plain)
            with zipfile.ZipFile(plain) as src, zipfile.ZipFile(path, "w") as out:
                for item in src.infolist():
                    data = src.read(item.filename)
                    if item.filename == "xl/worksheets/sheet1.xml":
                        data = data.replace(b"</worksheet>", X14_EXT.encode() + b"</worksheet>")
                    out.writestr(item, data)

            # Отдельный процесс, а не catch_warnings: раннер unittest ставит
            # свой фильтр «default» впереди всех и перебил бы наш. Консоль
            # запускается без раннера — так и проверяем.
            def run(prelude: str) -> str:
                code = ("import sys; sys.path.insert(0, %r); %s"
                        "import openpyxl; print(openpyxl.load_workbook(%r).active['A1'].value)"
                        % (str(BASE_DIR), prelude, str(path)))
                done = subprocess.run([sys.executable, "-c", code], capture_output=True,
                                      text=True, check=True)
                self.assertEqual(done.stdout.strip(), "1")
                return done.stderr

            # Без фильтра предупреждение есть — значит файл его действительно вызывает.
            self.assertIn("Data Validation extension", run(""))
            self.assertEqual(run("import common; "), "")


if __name__ == "__main__":
    unittest.main()
