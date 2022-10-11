from PyQt5 import QtCore, QtGui, QtWidgets
from technical_functions import *

from xlsx_data import XlsxData

from cell import Cell
from parser import Parser


class Excel(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.external_table = XlsxData()
        self.msg = MessageBox()
        self.cell = None
        self.rowCount = 4
        self.colCount = 4
        self.setObjectName("Excel")
        self.resize(867, 488)
        self.init_ui()

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
        self.pushButtonCalculate.setGeometry(QtCore.QRect(740, 30, 101, 31))
        self.pushButtonCalculate.setObjectName("pushButtonCalculate")

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
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 21))
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

        self.pushButtonAddRow.clicked.connect(self.row_btn_add)
        self.pushButtonDelRow.clicked.connect(self.row_btn_del)
        self.pushButtonAddCol.clicked.connect(self.col_btn_add)
        self.pushButtonDelCol.clicked.connect(self.col_btn_del)
        self.pushButtonCalculate.clicked.connect(self.input_line)

        self.actionOpen.triggered.connect(self.load_data)
        self.actionSave.triggered.connect(self.external_table.save_table)
        self.actionClear.triggered.connect(self.clear_table)
        self.actionAbout.triggered.connect(self.msg.about_project)

        self.tableWidget.selectionModel().selectionChanged.connect(self.get_selected_cell)
        self.tableWidget.itemChanged.connect(self.trace_changes)

    def row_btn_add(self):
        self.tableWidget.setRowCount(self.rowCount + 1)

        self.tableWidget.setVerticalHeaderItem(self.rowCount, QtWidgets.QTableWidgetItem(str(self.rowCount + 1)))

        self.external_table.row_xlsx_add(self.rowCount + 1)
        self.rowCount += 1

    def col_btn_add(self):
        base_char = 1
        for s in iter_all_strings():
            if base_char > self.colCount:
                self.tableWidget.setColumnCount(self.colCount + 1)

                self.tableWidget.setHorizontalHeaderItem(self.colCount, QtWidgets.QTableWidgetItem(s))

                self.external_table.col_xlsx_add(self.colCount + 1)
                base_char += 1
                self.colCount += 1
                break
            else:
                base_char += 1
                continue

    def row_btn_del(self):
        if self.rowCount > 3:
            self.tableWidget.setRowCount(self.rowCount - 1)
            self.external_table.row_xlsx_del(self.rowCount)
            self.rowCount -= 1
        else:
            self.msg.min_table_size_warning()

    def col_btn_del(self):
        if self.colCount > 3:
            self.tableWidget.setColumnCount(self.colCount - 1)
            self.external_table.col_xlsx_del(self.colCount)
            self.colCount -= 1
        else:
            self.msg.min_table_size_warning()

    def load_data(self):
        self.clear_table()
        self.external_table.reload_work_book()
        self.tableWidget.setRowCount(self.external_table.get_working_sheet().max_row)
        self.tableWidget.setColumnCount(self.external_table.get_working_sheet().max_column)
        self.set_horizontal_header_name()
        table_data = list(self.external_table.get_working_sheet().values)

        row_ix = 0
        for value_tuple in table_data:
            col_ix = 0
            for value in value_tuple:
                self.tableWidget.setItem(row_ix, col_ix, QtWidgets.QTableWidgetItem(str(value)))
                col_ix += 1
            row_ix += 1

        self.rowCount = self.tableWidget.rowCount()
        self.colCount = self.tableWidget.columnCount()

    def input_line(self):
        expression = self.lineEdit.text()
        if self.cell:
            if expression != '':
                self.cell.parsing(expression)
            else:
                self.msg.expression_field_is_empty()
        else:
            self.msg.cell_is_not_selected()

    def get_selected_cell(self, selected, deselected):
        for ix in selected.indexes():
            self.cell = Cell(ix.row(), ix.column(), self.tableWidget)

    def trace_changes(self):
        row = self.tableWidget.currentIndex().row()
        col = self.tableWidget.currentIndex().column()
        thing = self.tableWidget.item(row, col)
        if thing is not None and thing.text() != '':
            cell = Cell(row, col, self.tableWidget)
            if cell.get_cell_text(row, col)[0] == '=':
                cell.parsing(thing.text())

    def clear_table(self):
        for row in range(0, self.rowCount):
            for col in range(0, self.colCount):
                self.tableWidget.setItem(row, col, QtWidgets.QTableWidgetItem())
                col += 1
            row += 1

    def get_row_count(self):
        return self.rowCount

    def get_col_count(self):
        return self.colCount

    def set_horizontal_header_name(self):
        base_char = 1
        for s in iter_all_strings():
            if base_char > self.colCount:
                self.tableWidget.setHorizontalHeaderItem(self.colCount, QtWidgets.QTableWidgetItem(s))
                base_char += 1
                break
            else:
                base_char += 1
                continue

    def retranslate_ui(self, excel):
        _translate = QtCore.QCoreApplication.translate
        excel.setWindowTitle(_translate("Excel", "MeinLiebsterExcel"))

        self.pushButtonAddRow.setText(_translate("Excel", "Add Row"))
        self.pushButtonAddCol.setText(_translate("Excel", "Add Column"))
        self.pushButtonDelRow.setText(_translate("Excel", "Delete Row"))
        self.pushButtonDelCol.setText(_translate("Excel", "Delete Column"))
        self.pushButtonCalculate.setText(_translate("Excel", "Calculate"))

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
