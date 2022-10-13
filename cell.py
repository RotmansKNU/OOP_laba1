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
        self.parser = excel.Parser(expr, self.tableWidget)
        if expr[0] == '=':
            self.parsing_for_cell(expr)
        else:
            self.parsing_for_line(expr)

    def parsing_for_cell(self, expr):
        if expr[1:4] == 'max' or expr[1:4] == 'min':
            self.cell_comparing()
        else:
            self.cell_calculation()

    def parsing_for_line(self, expr):
        if expr[:3] == 'max' or expr[:3] == 'min':
            self.line_comparing()
        else:
            self.line_calculation()

    def line_calculation(self):
        res = self.parser.calculation_from_line()
        if res is not None:
            self.fill_cell(str(res))
        else:
            self.msg.incorrect_expression()

    def line_comparing(self):
        res = self.parser.comparing_from_line()
        if res is not None:
            self.fill_cell(str(res))
        else:
            self.msg.incorrect_expression()

    def cell_calculation(self):
        res = self.parser.calculation_from_cell()
        if res is not None and res is not False:
            self.fill_cell(str(res))
        elif res is False:
            self.fill_cell('')
        else:
            self.msg.incorrect_expression()

    def cell_comparing(self):
        res = self.parser.comparing_from_cell()
        if res is not None:
            self.fill_cell(str(res))
        else:
            self.msg.incorrect_expression()
