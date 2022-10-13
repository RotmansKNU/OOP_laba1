from PyQt5 import QtCore, QtGui, QtWidgets
import excel

import re
import operator


class Parser(excel.Cell):
    def __init__(self, expr, tableWidget):
        self.expression = expr
        self.tableWidget = tableWidget
        self.msg = excel.MessageBox()
        self.op = {'+': lambda x, y: x + y,
                   '-': lambda x, y: x - y,
                   '*': lambda x, y: x * y,
                   '/': lambda x, y: x / y if (y != 0) else self.msg.dividing_by_zero(),
                   '^': lambda x, y: x ** y,
                   'max': lambda x, y: max(x, y),
                   'min': lambda x, y: min(x, y)}

        # if calculate with none cell
        # replacing

    def calculation_from_cell(self):
        pattern = re.compile('^\=([A-Z]*)(\-\d*\.?\d*|\d*\.?\d*)(\+|\-|\*|\/|\^)([A-Z]*)(\-\d*\.?\d*|\d*\.?\d*)$')
        it = re.finditer(pattern, self.expression)
        operands = None
        first = True
        second = True
        parts = [0, 0]

        try:
            if it:
                for element in it:
                    operands = element.group(1, 2, 3, 4, 5)
                    if operands is not None:
                        for j in range(self.tableWidget.columnCount()):
                            if operands[0] == self.tableWidget.horizontalHeaderItem(j).text() and first:
                                parts[0] = self.tableWidget.item(int(operands[1]) - 1, j).text()
                                first = False
                            elif operands[3] == self.tableWidget.horizontalHeaderItem(j).text() and second:
                                parts[1] = self.tableWidget.item(int(operands[4]) - 1, j).text()
                                second = False

                            if operands[0] == '' and first:
                                parts[0] = operands[1]
                                first = False
                            if operands[3] == '' and second:
                                parts[1] = operands[4]
                                second = False

                        return self.op[operands[2]](float(parts[0]), float(parts[1]))
                    else:
                        return None
        except:
            self.msg.wrong_index()
            return False

    def comparing_from_cell(self):
        pattern = re.compile('^\=(\w{3})\(([A-Z]*)(\-\d*\.?\d*|\d*\.?\d*)\,\s?([A-Z]*)(\-\d*\.?\d*|\d*\.?\d*)\)$')
        it = re.finditer(pattern, self.expression)
        operands = None
        first = True
        second = True
        parts = [0, 0]

        try:
            if it:
                for element in it:
                    operands = element.group(1, 2, 3, 4, 5)
                    if operands is not None:
                        for j in range(self.tableWidget.columnCount()):
                            if operands[0] == self.tableWidget.horizontalHeaderItem(j).text() and first:
                                parts[0] = self.tableWidget.item(int(operands[1]) - 1, j).text()
                                first = False
                            elif operands[3] == self.tableWidget.horizontalHeaderItem(j).text() and second:
                                parts[1] = self.tableWidget.item(int(operands[4]) - 1, j).text()
                                second = False

                            if operands[0] == '' and first:
                                parts[0] = operands[1]
                                first = False
                            if operands[3] == '' and second:
                                parts[1] = operands[4]
                                second = False

                        return self.op[operands[0]](float(parts[0]), float(parts[1]))
                    else:
                        return None
        except:
            self.msg.wrong_index()
            return False

    def calculation_from_line(self):
        pattern = re.compile('^(\-\d*\.?\d*|\d*\.?\d*)\s?(\+|\-|\*|\/|\^)\s?(\-\d*\.?\d*|\d*\.?\d*)$')
        it = re.finditer(pattern, self.expression)
        operands = None
        if it:
            for element in it:
                operands = element.group(1, 2, 3)
                if operands is not None:
                    return self.op[operands[1]](float(operands[0]), float(operands[2]))
                else:
                    return None

    def comparing_from_line(self):
        pattern = re.compile('^(\w{3})\((\-\d*\.?\d*|\d*\.?\d*)\,\s?(\-\d*\.?\d*|\d*\.?\d*)\)$')
        it = re.finditer(pattern, self.expression)
        operands = None
        if it:
            for element in it:
                operands = element.group(1, 2, 3)
                if operands is not None:
                    return self.op[operands[0]](float(operands[1]), float(operands[2]))
                else:
                    return None
