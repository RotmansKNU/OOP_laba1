import Excel.excel as excel
import sys

from src.global_enums import GlobalErrorMessages


def test_row_btn_add():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()

    window.row_btn_add()

    cellB5 = excel.Cell(4, 1, window.get_table_widget())

    assert window.rowCount == 5, GlobalErrorMessages.AddingRowError.value
    assert window.get_row_count() == 5, GlobalErrorMessages.AddingRowError.value
    assert window.tableWidget.rowCount() == 5, GlobalErrorMessages.AddingRowError.value
    assert cellB5.get_cell_text(4, 1) == '', GlobalErrorMessages.AddingRowError.value
    assert window.tableWidget.verticalHeaderItem(4).text() == '5', GlobalErrorMessages.AddingRowError.value


def test_col_btn_add():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()

    window.col_btn_add()

    cellE1 = excel.Cell(0, 4, window.get_table_widget())

    assert window.colCount == 5, GlobalErrorMessages.AddingColError.value
    assert window.get_col_count() == 5, GlobalErrorMessages.AddingColError.value
    assert window.tableWidget.columnCount() == 5, GlobalErrorMessages.AddingColError.value
    assert cellE1.get_cell_text(0, 4) == '', GlobalErrorMessages.AddingColError.value
    assert window.tableWidget.horizontalHeaderItem(4).text() == 'E', GlobalErrorMessages.AddingColError.value


def test_row_btn_del():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()

    window.row_btn_del()

    assert window.rowCount == 3, GlobalErrorMessages.DeletingRowError.value
    assert window.get_row_count() == 3, GlobalErrorMessages.DeletingRowError.value
    assert window.tableWidget.rowCount() == 3, GlobalErrorMessages.DeletingRowError.value

    window.row_btn_del()
    window.row_btn_del()
    window.row_btn_del()

    assert window.rowCount == 1, GlobalErrorMessages.DeletingRowError.value


def test_col_btn_del():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()

    window.col_btn_del()

    cellB5 = excel.Cell(4, 1, window.get_table_widget())

    assert window.colCount == 3, GlobalErrorMessages.DeletingColError.value
    assert window.get_col_count() == 3, GlobalErrorMessages.DeletingColError.value
    assert window.tableWidget.columnCount() == 3, GlobalErrorMessages.DeletingColError.value

    window.col_btn_del()
    window.col_btn_del()
    window.col_btn_del()

    assert window.colCount == 1, GlobalErrorMessages.DeletingColError.value


def test_fill_line():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()

    window.fill_line()

    assert window.lineEdit.text() == '', GlobalErrorMessages.FillLineError.value


def test_clear_line():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()

    window.clear_line()

    assert window.lineEdit.text() == '', GlobalErrorMessages.ClearLineError.value


def test_load_data():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()

    window.is_saved = False

    window.load_data()

    assert window.is_saved is True, GlobalErrorMessages.ClearLineError.value


def test_add_text_to_line_on_stack():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()

    assert window.add_text_to_line_on_stack('world') == 'world', GlobalErrorMessages.ClearLineError.value
