import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from excel import Excel

if __name__ == "__main__":
    application = QtWidgets.QApplication(sys.argv)
    window = Excel()
    window.show()
    sys.exit(application.exec_())
