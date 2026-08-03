import pytest
from openpyxl import Workbook

from config.constants import EXCEL_SHEET_NAME
from excel.reader import ExcelReader


def _write_workbook(path, sheet_name=EXCEL_SHEET_NAME):
    workbook = Workbook()
    workbook.active.title = sheet_name
    workbook.save(path)


class TestRead:
    def test_reads_workbook_and_expected_sheet(self, tmp_path):
        path = tmp_path / "Bar_Piscina.xlsx"
        _write_workbook(path)

        reader = ExcelReader()
        workbook, worksheet = reader.read(workbook_path=path)

        assert worksheet.title == EXCEL_SHEET_NAME
        workbook.close()

    def test_missing_file_raises_file_not_found(self, tmp_path):
        reader = ExcelReader()

        with pytest.raises(FileNotFoundError):
            reader.read(workbook_path=tmp_path / "no_existe.xlsx")

    def test_directory_path_raises_value_error(self, tmp_path):
        reader = ExcelReader()

        with pytest.raises(ValueError):
            reader.read(workbook_path=tmp_path)

    def test_non_xlsx_extension_raises_value_error(self, tmp_path):
        path = tmp_path / "archivo.txt"
        path.write_text("no es un excel")

        reader = ExcelReader()

        with pytest.raises(ValueError):
            reader.read(workbook_path=path)

    def test_missing_expected_sheet_raises_value_error(self, tmp_path):
        path = tmp_path / "SinHoja.xlsx"
        _write_workbook(path, sheet_name="OtraHoja")

        reader = ExcelReader()

        with pytest.raises(ValueError, match=EXCEL_SHEET_NAME):
            reader.read(workbook_path=path)
