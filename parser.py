from PyQt5 import QtCore, QtGui, QtWidgets
import excel

import re
import operator

class Parser(excel.Cell):
    def __init__(self, expr, tableWidget):
        self.expression = expr
        self.tableWidget = tableWidget
        self.op = {'+': lambda x, y: x + y,
                   '-': lambda x, y: x - y,
                   '*': lambda x, y: x * y,
                   '/': lambda x, y: x / y,
                   '^': lambda x, y: x ** y,
                   'max': lambda x, y: max(x, y),
                   'min': lambda x, y: min(x, y)}

    def calculation_from_cell(self):
        print('cell')
        pattern = re.compile('^(\=[A-Z]{1,2}\d+)(\+|\-|\*|\/|\^)([A-Z]{1,2}\d+)$')
        if re.search(pattern, self.expression):
            print('letter')

    def calculation_from_line(self):
        pattern = re.compile('^(\-\d+|\d+)\s{0,1}(\+|\-|\*|\/|\^)\s{0,1}(\-\d+|\d+)$')
        it = re.finditer(pattern, self.expression)
        operands = None
        if it:
            for element in it:
                operands = element.group(1, 2, 3)
                if operands is not None:
                    return self.op[operands[1]](int(operands[0]), int(operands[2]))
                else:
                    return None

    def comparing_functions(self):
        pattern = re.compile('^(\w{3})\((\-\d+|\d+)\,\s{0,1}(\-\d+|\d+)\)$')
        it = re.finditer(pattern, self.expression)
        operands = None
        if it:
            for element in it:
                operands = element.group(1, 2, 3)
                if operands is not None:
                    return self.op[operands[0]](int(operands[1]), int(operands[2]))
                else:
                    return None
