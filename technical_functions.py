from PyQt5 import QtCore, QtGui, QtWidgets

from string import ascii_uppercase
import itertools


def min_table_size_warning():
    msg = QtWidgets.QMessageBox()
    msg.setIcon(QtWidgets.QMessageBox.Warning)
    msg.setWindowTitle("Warning")
    msg.setText("You can't use table less than 3x3 dimension")
    msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
    msg.exec_()


def about_project():
    msg = QtWidgets.QMessageBox()
    msg.setIcon(QtWidgets.QMessageBox.Information)
    msg.setWindowTitle("About project")
    msg.setText("Here you can publish your info")
    msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
    msg.exec_()


def cell_is_not_selected():
    msg = QtWidgets.QMessageBox()
    msg.setIcon(QtWidgets.QMessageBox.Warning)
    msg.setWindowTitle("Warning")
    msg.setText("Select the cell and press button \"Calculate\"")
    msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
    msg.exec_()


def expression_field_is_empty():
    msg = QtWidgets.QMessageBox()
    msg.setIcon(QtWidgets.QMessageBox.Warning)
    msg.setWindowTitle("Warning")
    msg.setText("Write expression and press button \"Calculate\"")
    msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
    msg.exec_()


def iter_all_strings():
    for size in itertools.count(1):
        for s in itertools.product(ascii_uppercase, repeat=size):
            yield "".join(s)

