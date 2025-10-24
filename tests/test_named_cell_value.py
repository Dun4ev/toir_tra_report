from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

from toir_tra_report_v1 import set_named_cell_value


def test_set_named_cell_value_updates_target_cell():
    workbook = Workbook()
    worksheet = workbook.active
    defined_name = DefinedName(name="pripmem", attr_text=f"'{worksheet.title}'!$I$22")
    workbook.defined_names[defined_name.name] = defined_name

    result = set_named_cell_value(workbook, "pripmem", "Sample Sender")

    assert result is True
    assert worksheet["I22"].value == "Sample Sender"


def test_set_named_cell_value_missing_defined_name_returns_false():
    workbook = Workbook()

    result = set_named_cell_value(workbook, "missing", "Sample Sender")

    assert result is False
