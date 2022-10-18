import Excel.excel as excel
import sys

from src.global_enums import GlobalErrorMessages


def test_set_path():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()
    pathTupple = ['D:/Programs/Pycharm/PyProjects/OOP_laba1/test.xlsx', 'Excel File (*.xlsx *.xls)']
    window.external_table.set_path(pathTupple)

    assert window.external_table.path == 'D:/Programs/Pycharm/PyProjects/OOP_laba1/test.xlsx', GlobalErrorMessages.SetPathError.value


def test_get_working_sheet():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()
    pathTupple = ['D:/Programs/Pycharm/PyProjects/OOP_laba1/test.xlsx', 'Excel File (*.xlsx *.xls)']
    window.external_table.set_path(pathTupple)

    window.external_table.reload_work_book()
    assert str(window.external_table.get_working_sheet()) == '<Worksheet "test">', GlobalErrorMessages.GetWorkingSheetError.value
