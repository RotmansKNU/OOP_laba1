from PyQt5 import QtCore, QtGui, QtWidgets


def min_table_size_warning():
    msg = QtWidgets.QMessageBox()
    msg.setIcon(QtWidgets.QMessageBox.Warning)
    msg.setWindowTitle("Warning")
    msg.setText("You can't use table less than 3x3 dimension")
    msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
    msg.exec_()


def max_column_size_warning():
    msg = QtWidgets.QMessageBox()
    msg.setIcon(QtWidgets.QMessageBox.Warning)
    msg.setWindowTitle("Warning")
    msg.setText("You can't use table wider than 26 columns")
    msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
    msg.exec_()


def about_project():
    msg = QtWidgets.QMessageBox()
    msg.setIcon(QtWidgets.QMessageBox.Information)
    msg.setWindowTitle("About project")
    msg.setText("Here you can publish your info")
    msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
    msg.exec_()
