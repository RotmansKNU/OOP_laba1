from PyQt5 import QtCore, QtGui, QtWidgets

from string import ascii_uppercase
import itertools

import os


class MessageBox:
    def __init__(self):
        self.msg = QtWidgets.QMessageBox()

    def min_table_col_warning(self):
        self.msg.setIcon(QtWidgets.QMessageBox.Warning)
        self.msg.setWindowTitle("Warning")
        self.msg.setText("You can't use table less than 1 column")
        self.msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        self.msg.exec_()

    def min_table_row_warning(self):
        self.msg.setIcon(QtWidgets.QMessageBox.Warning)
        self.msg.setWindowTitle("Warning")
        self.msg.setText("You can't use table less than 1 row")
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

    def save_before_close(self):
        self.msg.setIcon(QtWidgets.QMessageBox.Warning)
        self.msg.setWindowTitle("Warning")
        self.msg.setText("Your changes won't save!")
        self.msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        self.msg.exec_()

    def open_file(self, parent):
        return QtWidgets.QFileDialog.getOpenFileName(parent, 'Select a file', os.getcwd(), 'Excel File (*.xlsx *.xls)', 'Excel File (*.xlsx *.xls)')

    def save_file(self, parent):
        return QtWidgets.QFileDialog.getSaveFileName(parent, 'Select a file', '', 'Excel File (*.xlsx *.xls)', 'Excel File (*.xlsx *.xls)')

    def wrong_file_format(self):
        self.msg.setIcon(QtWidgets.QMessageBox.Critical)
        self.msg.setWindowTitle("Error")
        self.msg.setText("Incorrect file format!")
        self.msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        self.msg.exec_()


def iter_all_strings():
    for size in itertools.count(1):
        for s in itertools.product(ascii_uppercase, repeat=size):
            yield "".join(s)

