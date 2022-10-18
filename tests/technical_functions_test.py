import Excel.excel as excel
import sys

from src.global_enums import GlobalErrorMessages


def test_about_project():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()
    msg = excel.MessageBox()

    assert window.msg.about_project() == msg.about_project(), GlobalErrorMessages.AboutProjectMsgError.value


def test_cell_is_not_selected():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()
    msg = excel.MessageBox()

    assert window.msg.cell_is_not_selected() == msg.cell_is_not_selected(), GlobalErrorMessages.CellIsNotSelectedMsgError.value


def test_expression_field_is_empty():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()
    msg = excel.MessageBox()

    assert window.msg.expression_field_is_empty() == msg.expression_field_is_empty(), GlobalErrorMessages.ExpressionFieldIsEmptyMsgError.value


def test_wrong_file_format():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()
    msg = excel.MessageBox()

    assert window.msg.wrong_file_format() == msg.wrong_file_format(), GlobalErrorMessages.WrongFileFormatMsgError.value
