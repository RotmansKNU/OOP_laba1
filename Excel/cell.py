from PyQt5 import QtWidgets
import Excel.excel as excel


class Cell:
    def __init__(self, row, col, tableWidget):
        self.row = row
        self.col = col
        self.tableWidget = tableWidget
        self.msg = excel.MessageBox()
        self.dependencies = []

    def get_cell_text(self, row, col):
        return self.tableWidget.item(row, col).text()

    def fill_cell(self, expr):
        self.tableWidget.setItem(self.row, self.col, QtWidgets.QTableWidgetItem(expr))

    def append_dependencies(self, value):
        self.dependencies.append(value)

    def update_dependencies(self, value):
        for i in self.dependencies:
            self.dependencies[i] = value

    def get_dependencies(self):
        return self.dependencies

    def parsing(self, expr):
        self.expression = expr
        self.parser = excel.Parser(self.expression, self.tableWidget)
        if self.expression[0] == '=':
            self.parsing_for_cell()
        elif self.expression[0] == '#':
            return self.parsing_for_replacement()
        else:
            self.parsing_for_line()

    def parsing_for_cell(self):
        if self.expression[1:4] == 'max' or self.expression[1:4] == 'min':
            self.cell_comparing()
        else:
            self.cell_calculation()

    def parsing_for_line(self):
        if self.expression[:3] == 'max' or self.expression[:3] == 'min':
            self.line_comparing()
        else:
            self.line_calculation()

    def parsing_for_replacement(self):
        res = self.parser.replacement()
        if res is not None:
            self.fill_cell(str(res))
            return str(res)
        else:
            self.fill_cell(str(self.expression[1:]))

    def line_calculation(self):
        res = self.parser.calculation_from_line()
        if res is not None:
            self.fill_cell(str(res))

    def line_comparing(self):
        res = self.parser.comparing_from_line()
        if res is not None:
            self.fill_cell(str(res))

    def cell_calculation(self):
        res = self.parser.calculation_from_cell()
        if res is not None:
            self.fill_cell(str(res))
        else:
            self.fill_cell(str(self.expression[1:]))

    def cell_comparing(self):
        res = self.parser.comparing_from_cell()
        if res is not None:
            self.fill_cell(str(res))
        else:
            self.fill_cell(str(self.expression[1:]))
