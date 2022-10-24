from PyQt5 import QtCore, QtGui, QtWidgets
from Excel.technical_functions import *

from Excel.xlsx_data import XlsxData

from Excel.cell import Cell
from Excel.parser import Parser

import pandas as pd


class Excel(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.tableWidget = None
        self.is_saved = True
        self.chainLink = 0
        self.cellsDict = dict()
        self.msg = MessageBox()
        self.external_table = XlsxData()
        self.cell = None
        self.textInInputLine = ''
        self.rowCount = 4
        self.colCount = 4
        self.setObjectName("Excel")
        self.setMinimumSize(860, 490)
        self.setMaximumSize(860, 490)
        self.init_ui()
        self.on_event()
        self.clear_table()

    def init_ui(self):
        self.centralWidget = QtWidgets.QWidget(self)
        self.centralWidget.setObjectName("centralWidget")

        self.pushButtonAddRow = QtWidgets.QPushButton(self.centralWidget)
        self.pushButtonAddRow.setGeometry(QtCore.QRect(20, 20, 101, 31))
        self.pushButtonAddRow.setObjectName("pushButtonAddRow")

        self.pushButtonAddCol = QtWidgets.QPushButton(self.centralWidget)
        self.pushButtonAddCol.setGeometry(QtCore.QRect(150, 20, 101, 31))
        self.pushButtonAddCol.setObjectName("pushButtonAddCol")

        self.pushButtonDelRow = QtWidgets.QPushButton(self.centralWidget)
        self.pushButtonDelRow.setGeometry(QtCore.QRect(280, 20, 101, 31))
        self.pushButtonDelRow.setObjectName("pushButtonDelRow")

        self.pushButtonDelCol = QtWidgets.QPushButton(self.centralWidget)
        self.pushButtonDelCol.setGeometry(QtCore.QRect(410, 20, 101, 31))
        self.pushButtonDelCol.setObjectName("pushButtonDelCol")

        self.pushButtonCalculate = QtWidgets.QPushButton(self.centralWidget)
        self.pushButtonCalculate.setGeometry(QtCore.QRect(740, 10, 101, 31))
        self.pushButtonCalculate.setObjectName("pushButtonCalculate")

        self.pushButtonAC = QtWidgets.QPushButton(self.centralWidget)
        self.pushButtonAC.setGeometry(QtCore.QRect(740, 50, 101, 31))
        self.pushButtonAC.setObjectName("pushButtonAC")

        self.lineEdit = QtWidgets.QLineEdit(self.centralWidget)
        self.lineEdit.setGeometry(QtCore.QRect(530, 30, 200, 31))
        self.lineEdit.setObjectName("lineEdit")

        self.tableWidget = QtWidgets.QTableWidget(self.centralWidget)
        self.tableWidget.setGeometry(QtCore.QRect(20, 90, 821, 351))
        self.tableWidget.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.tableWidget.setLineWidth(1)
        self.tableWidget.setMidLineWidth(0)
        self.tableWidget.setShowGrid(True)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(4)
        self.tableWidget.setRowCount(4)

        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(3, item)

        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(3, item)

        self.setCentralWidget(self.centralWidget)

        self.menubar = QtWidgets.QMenuBar(self)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1080, 720))
        self.menubar.setObjectName("menubar")

        self.menuFile = QtWidgets.QMenu(self.menubar)
        self.menuFile.setObjectName("menuFile")

        self.actionOpen = QtWidgets.QAction(self)
        self.actionOpen.setObjectName("actionOpen")
        self.menuFile.addAction(self.actionOpen)

        self.actionSave = QtWidgets.QAction(self)
        self.actionSave.setObjectName("actionSave")
        self.menuFile.addAction(self.actionSave)

        self.actionClear = QtWidgets.QAction(self)
        self.actionClear.setObjectName("actionClear")
        self.menuFile.addAction(self.actionClear)

        self.actionAbout = QtWidgets.QAction(self)
        self.actionAbout.setObjectName("actionAbout")
        self.menuFile.addAction(self.actionAbout)

        self.menubar.addAction(self.menuFile.menuAction())
        self.setMenuBar(self.menubar)

        self.retranslate_ui(self)
        QtCore.QMetaObject.connectSlotsByName(self)

    def on_event(self):
        self.pushButtonAddRow.clicked.connect(self.row_btn_add)
        self.pushButtonDelRow.clicked.connect(self.row_btn_del)
        self.pushButtonAddCol.clicked.connect(self.col_btn_add)
        self.pushButtonDelCol.clicked.connect(self.col_btn_del)
        self.pushButtonCalculate.clicked.connect(self.calculate)
        self.pushButtonAC.clicked.connect(self.clear_line)

        self.actionOpen.triggered.connect(self.open_data)
        self.actionSave.triggered.connect(self.save_data)
        self.actionClear.triggered.connect(self.clear_table)
        self.actionAbout.triggered.connect(self.msg.about_project)

        self.tableWidget.selectionModel().selectionChanged.connect(self.get_selected_cell)
        self.tableWidget.itemChanged.connect(self.trace_changes)

    def closeEvent(self, event):
        if self.is_saved is not True:
            self.msg.save_before_close(event, self.tableWidget, self.save_data)

    def row_btn_add(self):
        self.tableWidget.setRowCount(self.rowCount + 1)

        self.tableWidget.setVerticalHeaderItem(self.rowCount, QtWidgets.QTableWidgetItem(str(self.rowCount + 1)))

        for j in range(self.colCount):
            self.tableWidget.setItem(self.rowCount, j, QtWidgets.QTableWidgetItem(''))

        self.rowCount += 1

    def col_btn_add(self):
        base_char = 1
        for s in iter_all_strings():
            if base_char > self.colCount:
                self.tableWidget.setColumnCount(self.colCount + 1)

                self.tableWidget.setHorizontalHeaderItem(self.colCount, QtWidgets.QTableWidgetItem(s))

                base_char += 1
                self.colCount += 1
                break
            else:
                base_char += 1
                continue

        for j in range(self.rowCount):
            self.tableWidget.setItem(j, self.colCount - 1, QtWidgets.QTableWidgetItem(''))

    def row_btn_del(self):
        self.is_saved = False
        if self.rowCount > 1:
            for j in range(self.colCount):
                try:
                    cellIDString = f'{self.tableWidget.horizontalHeaderItem(j).text()}{self.rowCount}'
                    dependenciesTupple = self.cellsDict[cellIDString].get_dependencies()
                    i = 0
                    while i < len(dependenciesTupple):
                        el = dependenciesTupple[i]
                        self.tableWidget.setItem(el[0], el[1], QtWidgets.QTableWidgetItem('Empty'))
                        i += 1
                except:
                    pass

            self.tableWidget.setRowCount(self.rowCount - 1)
            self.rowCount -= 1
        else:
            self.msg.min_table_row_warning()

    def col_btn_del(self):
        self.is_saved = False
        if self.colCount > 1:
            for j in range(self.rowCount):
                try:
                    cellIDString = f'{self.tableWidget.horizontalHeaderItem(self.colCount - 1).text()}{j + 1}'
                    dependenciesTupple = self.cellsDict[cellIDString].get_dependencies()
                    i = 0
                    while i < len(dependenciesTupple):
                        el = dependenciesTupple[i]
                        self.tableWidget.setItem(el[0], el[1], QtWidgets.QTableWidgetItem('Empty'))
                        i += 1
                except:
                    pass

            self.tableWidget.setColumnCount(self.colCount - 1)
            self.colCount -= 1
        else:
            self.msg.min_table_col_warning()

    def open_data(self):
        if self.is_saved:
            self.load_data()
        else:
            if self.msg.save_when_reopen(self.tableWidget, self.save_data):
                self.load_data()

    def load_data(self):
        try:
            self.external_table.set_path(self.msg.open_file(self))
            self.external_table.reload_work_book()
            self.clear_table()

            table_data = list(self.external_table.get_working_sheet().values)
            self.creating_sheet_for_data()

            row_ix = 0
            for value_tuple in table_data:
                col_ix = 0
                for value in value_tuple:
                    if value is not None:
                        self.tableWidget.setItem(row_ix, col_ix, QtWidgets.QTableWidgetItem(str(value)))
                    else:
                        self.tableWidget.setItem(row_ix, col_ix, QtWidgets.QTableWidgetItem(''))
                    col_ix += 1
                row_ix += 1

            self.rowCount = self.tableWidget.rowCount()
            self.colCount = self.tableWidget.columnCount()
            self.is_saved = True
        except:
            self.msg.wrong_file_format()

    def save_data(self):
        try:
            path = self.msg.save_file(self)

            columnHeaders = []
            for j in range(self.tableWidget.model().columnCount()):
                columnHeaders.append(self.tableWidget.horizontalHeaderItem(j).text())

            df = pd.DataFrame(columns=columnHeaders)

            for row in range(self.tableWidget.rowCount()):
                for col in range(self.tableWidget.columnCount()):
                    df.at[row, columnHeaders[col]] = self.tableWidget.item(row, col).text()

            df.to_excel(path[0], header=False, index=False)
            self.is_saved = True
        except:
            self.msg.wrong_file_format()

    def creating_sheet_for_data(self):
        maxRow = self.external_table.get_working_sheet().max_row
        maxCol = self.external_table.get_working_sheet().max_column
        for it in range(1, maxRow):
            if self.rowCount < maxRow:
                self.row_btn_add()
        for it in range(1, maxCol):
            if self.colCount < maxCol:
                self.col_btn_add()

    def calculate(self):
        expression = self.lineEdit.text()
        if self.cell:
            if expression != '':
                self.cell.parsing(expression)
            else:
                self.msg.expression_field_is_empty()
        else:
            self.msg.cell_is_not_selected()

    def fill_line(self):
        row = self.tableWidget.currentIndex().row()
        col = self.tableWidget.currentIndex().column()
        thing = self.tableWidget.item(row, col)
        if thing is not None and thing.text() != '' and thing.text()[0] != '=' and thing.text()[0] != '#':
            self.lineEdit.setText(self.add_text_to_line_on_stack(thing.text()))

    def add_text_to_line_on_stack(self, txt):
        expression = self.lineEdit.text()
        self.textInInputLine = expression + txt
        return self.textInInputLine

    def get_selected_cell(self, selected, deselected):
        for ix in selected.indexes():
            self.cell = Cell(ix.row(), ix.column(), self.tableWidget)
            self.fill_line()

    def trace_changes(self):
        self.is_saved = False
        row = self.tableWidget.currentIndex().row()
        col = self.tableWidget.currentIndex().column()
        thing = self.tableWidget.item(row, col)

        if thing is not None and thing.text() != '':
            cell = Cell(row, col, self.tableWidget)
            cellID = [row, col]
            cellIDString = f'{self.tableWidget.horizontalHeaderItem(col).text()}{row + 1}'
            try:
                self.update_chain(cellIDString, cell, row, col, thing)
            except:
                self.create_chain(cellIDString, cellID, cell, row, col, thing)

    def update_chain(self, cellIDString, cell, row, col, thing):
        dependenciesTupple = self.cellsDict[cellIDString].get_dependencies()
        if self.chainLink >= len(dependenciesTupple) and len(dependenciesTupple) != 0:
            self.chainLink = 0
            return

        el = dependenciesTupple[self.chainLink]
        self.chainLink += 1
        if cell.get_cell_text(row, col)[0] == '=':
            res = cell.parsing(thing.text())
            self.tableWidget.setItem(el[0], el[1], QtWidgets.QTableWidgetItem(res))
        elif cell.get_cell_text(row, col)[0] == '#':
            res = cell.parsing(thing.text())
            self.tableWidget.setItem(el[0], el[1], QtWidgets.QTableWidgetItem(res))
        else:
            self.tableWidget.setItem(el[0], el[1], QtWidgets.QTableWidgetItem(thing.text()))

    def create_chain(self, cellIDString, cellID, cell, row, col, thing):
        if cell.get_cell_text(row, col)[0] == '#':
            baseID = cell.get_cell_text(row, col)[1:]
            try:
                self.cellsDict[baseID].append_dependencies(cellID)
                res = cell.parsing(thing.text())
                self.cellsDict[cellIDString] = res
            except:
                self.cellsDict[baseID] = cell
                self.cellsDict[baseID].append_dependencies(cellID)
                res = cell.parsing(thing.text())
                self.cellsDict[cellIDString] = res
        elif cell.get_cell_text(row, col)[0] == '=':
            cell.parsing(thing.text())
        else:
            self.cellsDict[cellIDString] = cell

    def clear_table(self):
        if self.is_saved:
            self.clear()
        else:
            if self.msg.save_when_clear(self.tableWidget, self.save_data):
                self.clear()

    def clear(self):
        for row in range(0, self.rowCount):
            for col in range(0, self.colCount):
                self.tableWidget.setItem(row, col, QtWidgets.QTableWidgetItem())
                col += 1
            row += 1
        self.is_saved = True

    def clear_line(self):
        self.lineEdit.setText('')

    def get_row_count(self):
        return self.rowCount

    def get_col_count(self):
        return self.colCount

    def get_table_widget(self):
        return self.tableWidget

    def retranslate_ui(self, excel):
        _translate = QtCore.QCoreApplication.translate
        excel.setWindowTitle(_translate("Excel", "MeinLiebsterExcel"))

        self.pushButtonAddRow.setText(_translate("Excel", "Add Row"))
        self.pushButtonAddCol.setText(_translate("Excel", "Add Column"))
        self.pushButtonDelRow.setText(_translate("Excel", "Delete Row"))
        self.pushButtonDelCol.setText(_translate("Excel", "Delete Column"))
        self.pushButtonCalculate.setText(_translate("Excel", "Calculate"))
        self.pushButtonAC.setText(_translate("Excel", "AC"))

        self.tableWidget.setSortingEnabled(False)

        item = self.tableWidget.verticalHeaderItem(0)
        item.setText(_translate("Excel", "1"))
        item = self.tableWidget.verticalHeaderItem(1)
        item.setText(_translate("Excel", "2"))
        item = self.tableWidget.verticalHeaderItem(2)
        item.setText(_translate("Excel", "3"))
        item = self.tableWidget.verticalHeaderItem(3)
        item.setText(_translate("Excel", "4"))

        item = self.tableWidget.horizontalHeaderItem(0)
        item.setText(_translate("Excel", "A"))
        item = self.tableWidget.horizontalHeaderItem(1)
        item.setText(_translate("Excel", "B"))
        item = self.tableWidget.horizontalHeaderItem(2)
        item.setText(_translate("Excel", "C"))
        item = self.tableWidget.horizontalHeaderItem(3)
        item.setText(_translate("Excel", "D"))

        self.menuFile.setTitle(_translate("Excel", "File"))
        self.actionOpen.setText(_translate("Excel", "Open"))
        self.actionSave.setText(_translate("Excel", "Save"))
        self.actionSave.setShortcut(_translate("Excel", "Ctrl+S"))
        self.actionClear.setText(_translate("Excel", "Clear"))
        self.actionAbout.setText(_translate("Excel", "About"))

        __sortingEnabled = self.tableWidget.isSortingEnabled()
        self.tableWidget.setSortingEnabled(False)
        self.tableWidget.setSortingEnabled(__sortingEnabled)
