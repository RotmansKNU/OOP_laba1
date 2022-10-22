from PyQt5 import QtWidgets
import Excel.excel as excel


class CellChain:
    def __init__(self, cell, data):
        self.pair = [cell, data]

    def on_changing(self, data):
        self.pair[1] = data

    def get_data(self):
        return self.pair[1]
