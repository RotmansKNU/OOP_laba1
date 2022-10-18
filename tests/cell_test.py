import Excel.excel as excel
import sys

from src.global_enums import GlobalErrorMessages


def test_get_cell_text():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()

    cellB1 = excel.Cell(0, 1, window.get_table_widget())
    cellB1.fill_cell(str(5))

    assert cellB1.get_cell_text(0, 1) == str(5), GlobalErrorMessages.GetCellTextError.value


def test_parsing():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()

    cellB1 = excel.Cell(0, 1, window.get_table_widget())
    cellB1.fill_cell('')
    cellB2 = excel.Cell(1, 1, window.get_table_widget())
    cellB2.fill_cell('')
    cellB3 = excel.Cell(2, 1, window.get_table_widget())
    cellB3.fill_cell('2')

    cellB1.parsing('3+5')
    cellB2.parsing('=-B3+5')
    cellB3.parsing('#C3')

    assert float(cellB1.get_cell_text(0, 1)) == 8.0, GlobalErrorMessages.ParsingError.value
    assert float(cellB1.get_cell_text(1, 1)) == 3.0, GlobalErrorMessages.ParsingError.value
    assert float(cellB1.get_cell_text(2, 1)) == 0.0, GlobalErrorMessages.ParsingError.value


def test_parsing_for_cell():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()

    cellB1 = excel.Cell(0, 1, window.get_table_widget())
    cellB1.fill_cell('')
    cellB1.parsing('=3+5')

    cellB2 = excel.Cell(1, 1, window.get_table_widget())
    cellB2.fill_cell('')
    cellB2.parsing('=max(-2, 5)')

    assert float(cellB1.get_cell_text(0, 1)) == 8.0, GlobalErrorMessages.CellParsingError.value
    assert float(cellB1.get_cell_text(1, 1)) == 5.0, GlobalErrorMessages.CellParsingError.value


def test_parsing_for_line():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()

    cellB1 = excel.Cell(0, 1, window.get_table_widget())
    cellB1.fill_cell('')
    cellB1.parsing('2^3')

    cellB2 = excel.Cell(1, 1, window.get_table_widget())
    cellB2.fill_cell('')
    cellB2.parsing('min(4, -0.4)')

    assert float(cellB1.get_cell_text(0, 1)) == 8.0, GlobalErrorMessages.LineParsingError.value
    assert float(cellB1.get_cell_text(1, 1)) == -0.4, GlobalErrorMessages.LineParsingError.value


def test_parsing_for_replacement():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()

    cellB1 = excel.Cell(0, 1, window.get_table_widget())
    cellB1.fill_cell('')
    cellB1.parsing('#G3')

    assert cellB1.get_cell_text(0, 1) == 'G3', GlobalErrorMessages.ReplacementParsingError.value


def test_cell_calculation():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()

    cellB1 = excel.Cell(0, 1, window.get_table_widget())
    cellB1.fill_cell('')
    cellB1.parsing('=G3+5')

    cellB2 = excel.Cell(1, 1, window.get_table_widget())
    cellB2.fill_cell('')
    cellB3 = excel.Cell(2, 1, window.get_table_widget())
    cellB3.fill_cell('')
    cellB2.parsing('=B3+5')

    assert cellB1.get_cell_text(0, 1) == 'G3+5', GlobalErrorMessages.CellCalculationError.value
    assert float(cellB2.get_cell_text(1, 1)) == 5.0, GlobalErrorMessages.CellCalculationError.value


def test_cell_comparing():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()

    cellB1 = excel.Cell(0, 1, window.get_table_widget())
    cellB1.fill_cell('')

    cellB2 = excel.Cell(1, 1, window.get_table_widget())
    cellB2.fill_cell('-5.4')

    cellB3 = excel.Cell(2, 1, window.get_table_widget())
    cellB3.fill_cell('')
    cellB3.parsing('=max(B1, B2)')

    assert float(cellB3.get_cell_text(2, 1)) == 0, GlobalErrorMessages.CellComparingError.value


def test_cell_comparing_exception():
    application = excel.QtWidgets.QApplication(sys.argv)
    window = excel.Excel()

    cellB1 = excel.Cell(0, 1, window.get_table_widget())
    cellB1.fill_cell('-5.4')

    cellB2 = excel.Cell(1, 1, window.get_table_widget())
    cellB2.fill_cell('')
    cellB2.parsing('=max(B4, B5)')

    assert cellB2.get_cell_text(1, 1) == 'max(B4, B5)', GlobalErrorMessages.CellComparingError.value
