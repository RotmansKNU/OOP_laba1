import sys
from PyQt5 import QtCore, QtGui, QtWidgets
import Excel.excel as excel

if __name__ == "__main__":
    application = QtWidgets.QApplication(sys.argv)
    window = excel.Excel()
    window.show()
    sys.exit(application.exec_())
