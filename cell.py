from PyQt5 import QtCore, QtGui, QtWidgets
import excel


class Cell:
    def __init__(self, row, col, tableWidget):
        self.row = row
        self.col = col
        self.tableWidget = tableWidget
        self.msg = excel.MessageBox()

    def get_cell_text(self, row, col):
        return self.tableWidget.item(row, col).text()

    def fill_cell(self, expr):
        self.tableWidget.setItem(self.row, self.col, QtWidgets.QTableWidgetItem(expr))

    def parsing(self, expr):
        parser = excel.Parser(expr, self.tableWidget)
        if expr[0] == '=':
            res = parser.calculation_from_cell()
            if res is not None:
                self.fill_cell(str(res))
            else:
                self.msg.incorrect_expression()
        elif expr[:3] == 'max' or expr[:3] == 'min':
            res = parser.comparing_functions()
            if res is not None:
                self.fill_cell(str(res))
            else:
                self.msg.incorrect_expression()
        else:
            res = parser.calculation_from_line()
            if res is not None:
                self.fill_cell(str(res))
            else:
                self.msg.incorrect_expression()
