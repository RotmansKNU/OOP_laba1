import Excel.excel as excel

import re


class Parser:
    def __init__(self, expr, tableWidget):
        self.expression = expr
        self.tableWidget = tableWidget
        self.msg = excel.MessageBox()
        self.op = {'+': lambda x, y: x + y,
                   '-': lambda x, y: x - y,
                   '*': lambda x, y: x * y,
                   '/': lambda x, y: x / y if (y != 0) else self.msg.dividing_by_zero(x, y),
                   '^': lambda x, y: x ** y if (y != 0 or x != 0) else self.msg.zero_to_pover_of_zero(x, y),
                   'max': lambda x, y: max(x, y) if (x != y) else x,
                   'min': lambda x, y: min(x, y) if (x != y) else x}

    def calculation_from_cell(self):
        pattern = re.compile('^\=(\-?)([A-Z]*)(\-\d*\.?\d*|\d*\.?\d*)\s?(\+|\-|\*|\/|\^)\s?([A-Z]*)(\-\d*\.?\d*|\d*\.?\d*)$')
        it = re.finditer(pattern, self.expression)
        operands = None
        first = True
        second = True
        parts = [None, None]

        try:
            for element in it:
                operands = element.group(1, 2, 3, 4, 5, 6)
                if operands is not None:
                    for j in range(self.tableWidget.columnCount()):
                        if operands[1] == self.tableWidget.horizontalHeaderItem(j).text() and first:
                            if self.tableWidget.item(int(operands[2]) - 1, j).text() != '' and self.tableWidget.item(int(operands[2]) - 1, j).text() != self.expression:
                                parts[0] = self.tableWidget.item(int(operands[2]) - 1, j).text()
                                first = False
                            else:
                                parts[0] = 0
                                first = False
                        if operands[4] == self.tableWidget.horizontalHeaderItem(j).text() and second:
                            if self.tableWidget.item(int(operands[5]) - 1, j).text() != '' and self.tableWidget.item(int(operands[5]) - 1, j).text() != self.expression:
                                parts[1] = self.tableWidget.item(int(operands[5]) - 1, j).text()
                                second = False
                            else:
                                parts[1] = 0
                                second = False

                        if operands[1] == '' and first:
                            parts[0] = operands[2]
                            first = False
                        if operands[4] == '' and second:
                            parts[1] = operands[5]
                            second = False

                    if operands[0] == '-':
                        parts[0] = float(parts[0]) * -1
                    if parts[0] is not None and parts[1] is not None:
                        return self.op[operands[3]](float(parts[0]), float(parts[1]))
            raise
        except:
            self.msg.wrong_index()

    def comparing_from_cell(self):
        pattern = re.compile('^\=(\w{3})\(([A-Z]*)(\-\d*\.?\d*|\d*\.?\d*)\,\s?([A-Z]*)(\-\d*\.?\d*|\d*\.?\d*)\)$')
        it = re.finditer(pattern, self.expression)
        operands = None
        first = True
        second = True
        parts = [None, None]

        try:
            for element in it:
                operands = element.group(1, 2, 3, 4, 5)
                if operands is not None:
                    for j in range(self.tableWidget.columnCount()):
                        if operands[1] == self.tableWidget.horizontalHeaderItem(j).text() and first:
                            if self.tableWidget.item(int(operands[2]) - 1, j).text() != '' and self.tableWidget.item(int(operands[2]) - 1, j).text() != self.expression:
                                parts[0] = self.tableWidget.item(int(operands[2]) - 1, j).text()
                                first = False
                            else:
                                parts[0] = 0
                                first = False
                        if operands[3] == self.tableWidget.horizontalHeaderItem(j).text() and second:
                            if self.tableWidget.item(int(operands[4]) - 1, j).text() != '' and self.tableWidget.item(int(operands[4]) - 1, j).text() != self.expression:
                                parts[1] = self.tableWidget.item(int(operands[4]) - 1, j).text()
                                second = False
                            else:
                                parts[1] = 0
                                second = False

                        if operands[1] == '' and first:
                            parts[0] = operands[2]
                            first = False
                        if operands[3] == '' and second:
                            parts[1] = operands[4]
                            second = False

                    if parts[0] is not None and parts[1] is not None:
                        return self.op[operands[0]](float(parts[0]), float(parts[1]))
            raise
        except:
            self.msg.wrong_index()

    def calculation_from_line(self):
        pattern = re.compile('^(\-\d*\.?\d*|\d*\.?\d*)\s?(\+|\-|\*|\/|\^)\s?(\-\d*\.?\d*|\d*\.?\d*)$')
        it = re.finditer(pattern, self.expression)
        operands = None

        for element in it:
            operands = element.group(1, 2, 3)
            if operands is not None:
                return self.op[operands[1]](float(operands[0]), float(operands[2]))
        self.msg.incorrect_expression()

    def comparing_from_line(self):
        pattern = re.compile('^(\w{3})\((\-\d*\.?\d*|\d*\.?\d*)\,\s?(\-\d*\.?\d*|\d*\.?\d*)\)$')
        it = re.finditer(pattern, self.expression)
        operands = None

        for element in it:
            operands = element.group(1, 2, 3)
            if operands is not None:
                return self.op[operands[0]](float(operands[1]), float(operands[2]))
        self.msg.incorrect_expression()

    def replacement(self):
        pattern = re.compile('^\#([A-Z]+)(\d+)$')
        it = re.finditer(pattern, self.expression)
        operands = None

        try:
            for element in it:
                operands = element.group(1, 2)
                if operands is not None:
                    for j in range(self.tableWidget.columnCount()):
                        if operands[0] == self.tableWidget.horizontalHeaderItem(j).text():
                            if self.tableWidget.item(int(operands[1]) - 1, j).text() != '' and self.tableWidget.item(int(operands[1]) - 1, j).text() != self.expression:
                                return self.tableWidget.item(int(operands[1]) - 1, j).text()
                            else:
                                return 0
            raise
        except:
            self.msg.wrong_index()
