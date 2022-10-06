from PyQt5 import QtCore, QtGui, QtWidgets


class ExcelUi(QtWidgets.QMainWindow):
    def __init__(self):
        super(ExcelUi, self).__init__()
        self.rowCount = 4
        self.colCount = 4
        self.setObjectName("Excel")
        self.resize(867, 488)
        self.init_ui()

    def init_ui(self):
        self.centralWidget = QtWidgets.QWidget(self)
        self.centralWidget.setObjectName("centralWidget")

        self.pushButtonRow = QtWidgets.QPushButton(self.centralWidget)
        self.pushButtonRow.setGeometry(QtCore.QRect(20, 20, 101, 31))
        self.pushButtonRow.setObjectName("pushButtonRow")

        self.pushButtonCol = QtWidgets.QPushButton(self.centralWidget)
        self.pushButtonCol.setGeometry(QtCore.QRect(150, 20, 101, 31))
        self.pushButtonCol.setObjectName("pushButtonCol")

        self.pushButtonDelRow = QtWidgets.QPushButton(self.centralWidget)
        self.pushButtonDelRow.setGeometry(QtCore.QRect(280, 20, 101, 31))
        self.pushButtonDelRow.setObjectName("pushButtonDelRow")

        self.pushButtonDelCol = QtWidgets.QPushButton(self.centralWidget)
        self.pushButtonDelCol.setGeometry(QtCore.QRect(410, 20, 101, 31))
        self.pushButtonDelCol.setObjectName("pushButtonDelCol")

        self.lineEdit = QtWidgets.QLineEdit(self.centralWidget)
        self.lineEdit.setGeometry(QtCore.QRect(530, 30, 311, 31))
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

        #item = QtWidgets.QTableWidgetItem()
        #self.tableWidget.setItem(1, 1, item)
        #item = QtWidgets.QTableWidgetItem()
        #self.tableWidget.setItem(3, 3, item)
        self.setCentralWidget(self.centralWidget)

        #self.statusbar = QtWidgets.QStatusBar(excel)
        #self.statusbar.setObjectName("statusbar")
        #excel.setStatusBar(self.statusbar)
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

        self.pushButtonRow.clicked.connect(self.row_btn_clicked)
        self.pushButtonCol.clicked.connect(self.col_btn_clicked)

    def row_btn_clicked(self):
        self.tableWidget.setRowCount(self.rowCount + 1)

        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(self.rowCount, item)

        item = self.tableWidget.verticalHeaderItem(self.rowCount)
        item.setText(QtCore.QCoreApplication.translate("Excel", str(self.rowCount + 1)))

        self.rowCount += 1

    def col_btn_clicked(self):
        self.tableWidget.setColumnCount(self.colCount + 1)

        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(self.colCount, item)

        item = self.tableWidget.horizontalHeaderItem(self.colCount)
        item.setText(QtCore.QCoreApplication.translate("Excel", "E"))

        self.colCount += 1

    def retranslate_ui(self, excel):
        _translate = QtCore.QCoreApplication.translate
        excel.setWindowTitle(_translate("Excel", "MeinLiebsterExcel"))

        self.pushButtonRow.setText(_translate("Excel", "Add Row"))
        self.pushButtonCol.setText(_translate("Excel", "Add Column"))
        self.pushButtonDelRow.setText(_translate("Excel", "Delete Row"))
        self.pushButtonDelCol.setText(_translate("Excel", "Delete Column"))

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
