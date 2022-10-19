from PyQt5 import QtCore, QtGui, QtWidgets

from string import ascii_uppercase
import itertools

import os


class MessageBox:
    def __init__(self):
        self.msg = QtWidgets.QMessageBox()

    def min_table_col_warning(self):
        self.msg.setIcon(self.msg.Warning)
        self.msg.setWindowTitle("Warning")
        self.msg.setText("You can't use table less than 1 column")
        self.msg.setStandardButtons(self.msg.Ok)
        self.msg.exec_()

    def wrong_index(self):
        self.msg.setIcon(self.msg.Critical)
        self.msg.setWindowTitle("Error")
        self.msg.setText("You input wrong index!")
        self.msg.setStandardButtons(self.msg.Ok)
        self.msg.exec_()

    def min_table_row_warning(self):
        self.msg.setIcon(self.msg.Warning)
        self.msg.setWindowTitle("Warning")
        self.msg.setText("You can't use table less than 1 row")
        self.msg.setStandardButtons(self.msg.Ok)
        self.msg.exec_()

    def about_project(self):
        s = 'Цей додаток був створений задовго до появи оригінального Excel!\n' \
            'Тут можна виконати такі дії:\n' \
            '1) У лініїї введення:\n\t' \
            '1. Додавання\n\t' \
            '2. Віднімання\n\t' \
            '3. Множення\n\t' \
            '4. Ділення\n\t' \
            '5. Піднесення до степеня\n\t' \
            '6. Знаходження максимального та мінімального серед двох значень\n\t' \
            '7. Підтримуються унарні операціі\n' \
            '2) У комірках, за умови що перед виразом є \'=\':\n' \
            'Приклад 1: =B2 + 3\n\t' \
            '1. Всі операії перераховані у першому пункті\n\t' \
            '2. Внесення у клітинку значення іншої клітинки, за умови що перед виразом є \'#\'\n' \
            'Приклад 2: #B3\n' \
            '3) Відкрити існуючий файл з розширенням xlsx або xls\n' \
            '4) Зберігти таблицю у файл з розширенням xlsx або xls\n' \
            '5) Очистити всю таблицю\n'
        self.msg.setIcon(self.msg.Information)
        self.msg.setWindowTitle("About project")
        self.msg.setText(s)
        self.msg.setStandardButtons(self.msg.Ok)
        self.msg.exec_()

    def cell_is_not_selected(self):
        self.msg.setIcon(self.msg.Warning)
        self.msg.setWindowTitle("Warning")
        self.msg.setText("Select the cell!")
        self.msg.setStandardButtons(self.msg.Ok)
        self.msg.exec_()

    def expression_field_is_empty(self):
        self.msg.setIcon(self.msg.Warning)
        self.msg.setWindowTitle("Warning")
        self.msg.setText("Write expression and press button \"Calculate\"")
        self.msg.setStandardButtons(self.msg.Ok)
        self.msg.exec_()

    def incorrect_expression(self):
        self.msg.setIcon(self.msg.Critical)
        self.msg.setWindowTitle("Error")
        self.msg.setText("You input incorrect expression!")
        self.msg.setStandardButtons(self.msg.Ok)
        self.msg.exec_()

    def save_before_close(self, event, widget, saving_func):
        reply = self.msg.question(widget, 'Window Close', 'Do you want to close the window without saving?', self.msg.Yes | self.msg.Save | self.msg.Cancel)
        if reply == self.msg.Yes:
            event.accept()
        elif reply == self.msg.Save:
            saving_func()
            event.ignore()
        elif reply == self.msg.Cancel:
            event.ignore()

    def save_when_reopen(self, widget, saving_func):
        reply = self.msg.question(widget, 'Window Close', 'Do you want to open new file without saving previous?', self.msg.Yes | self.msg.Save | self.msg.Cancel)
        if reply == self.msg.Save:
            saving_func()
        elif reply == self.msg.Yes:
            return True
        elif reply == self.msg.Cancel:
            pass

    def save_when_clear(self, widget, saving_func):
        reply = self.msg.question(widget, 'Window Close', 'Do you want to clear table without saving?', self.msg.Yes | self.msg.Save | self.msg.Cancel)
        if reply == self.msg.Save:
            saving_func()
        elif reply == self.msg.Yes:
            return True
        elif reply == self.msg.Cancel:
            pass

    def open_file(self, parent):
        return QtWidgets.QFileDialog.getOpenFileName(parent, 'Select a file', os.getcwd(), 'Excel File (*.xlsx *.xls)', 'Excel File (*.xlsx *.xls)')

    def save_file(self, parent):
        return QtWidgets.QFileDialog.getSaveFileName(parent, 'Select a file', '', 'Excel File (*.xlsx *.xls)', 'Excel File (*.xlsx *.xls)')

    def wrong_file_format(self):
        self.msg.setIcon(self.msg.Critical)
        self.msg.setWindowTitle("Error")
        self.msg.setText("Incorrect file format!")
        self.msg.setStandardButtons(self.msg.Ok)
        self.msg.exec_()

    def dividing_by_zero(self, x, y):
        self.msg.setIcon(self.msg.Critical)
        self.msg.setWindowTitle("Error")
        self.msg.setText("You can't divide by zero!")
        self.msg.setStandardButtons(self.msg.Ok)
        self.msg.exec_()
        return f'{x}/{y}'

    def zero_to_pover_of_zero(self, x, y):
        self.msg.setIcon(self.msg.Critical)
        self.msg.setWindowTitle("Error")
        self.msg.setText("You can’t pover zero to zero!")
        self.msg.setStandardButtons(self.msg.Ok)
        self.msg.exec_()
        return f'{x}^{y}'


def iter_all_strings():
    for size in itertools.count(1):
        for s in itertools.product(ascii_uppercase, repeat=size):
            yield "".join(s)

