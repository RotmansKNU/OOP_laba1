from openpyxl import Workbook, load_workbook


class XlsxData:
    def __init__(self):
        self.path = None

    def get_working_sheet(self):
        return self.ws

    def set_path(self, path):
        self.path = path[0]

    def reload_work_book(self):
        self.wb = load_workbook(self.path)
        self.ws = self.wb.active
