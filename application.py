import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from excel_ui import ExcelUi

if __name__ == "__main__":
    application = QtWidgets.QApplication(sys.argv)
    window = ExcelUi()
    window.show()
    sys.exit(application.exec_())
