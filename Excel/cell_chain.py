from PyQt5 import QtWidgets
import Excel.excel as excel


class CellChain:
    def __init__(self, row, col, tableWidget):
        self.refCounter = 0
        self.row = row
        self.col = col
        self.tableWidget = tableWidget
        self.cell = excel.Cell(self.row, self.col, self.tableWidget)
        self.chain_node = None

    def inc_ref_count(self):
        self.refCounter += 1

    def dec_ref_count(self):
        self.refCounter -= 1

    def get_ref_count(self):
        return self.refCounter
