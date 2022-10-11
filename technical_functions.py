from PyQt5 import QtCore, QtGui, QtWidgets

from string import ascii_uppercase
import itertools


class MessageBox:
    def __init__(self):
        self.msg = QtWidgets.QMessageBox()

    def min_table_size_warning(self):
        self.msg.setIcon(QtWidgets.QMessageBox.Warning)
        self.msg.setWindowTitle("Warning")
        self.msg.setText("You can't use table less than 3x3 dimension")
        self.msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        self.msg.exec_()

    def about_project(self):
        self.msg.setIcon(QtWidgets.QMessageBox.Information)
        self.msg.setWindowTitle("About project")
        self.msg.setText("Here you can publish your info")
        self.msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        self.msg.exec_()

    def cell_is_not_selected(self):
        self.msg.setIcon(QtWidgets.QMessageBox.Warning)
        self.msg.setWindowTitle("Warning")
        self.msg.setText("Select the cell!")
        self.msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        self.msg.exec_()

    def expression_field_is_empty(self):
        self.msg.setIcon(QtWidgets.QMessageBox.Warning)
        self.msg.setWindowTitle("Warning")
        self.msg.setText("Write expression and press button \"Calculate\"")
        self.msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        self.msg.exec_()

    def incorrect_expression(self):
        self.msg.setIcon(QtWidgets.QMessageBox.Critical)
        self.msg.setWindowTitle("Error")
        self.msg.setText("You input incorrect expression!")
        self.msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        self.msg.exec_()


def iter_all_strings():
    for size in itertools.count(1):
        for s in itertools.product(ascii_uppercase, repeat=size):
            yield "".join(s)

