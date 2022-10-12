from openpyxl import Workbook, load_workbook


class XlsxData:
    def __init__(self):
        self.path = None

        #self.ws['A2'].value = 'col add'
    def row_xlsx_add(self, row):
        self.ws.insert_rows(row)

    def col_xlsx_add(self, col):
        self.ws.insert_cols(col)

    def row_xlsx_del(self, row):
        self.ws.delete_rows(row)

    def col_xlsx_del(self, col):
        self.ws.delete_cols(col)

    def get_working_sheet(self):
        return self.ws

    def set_path(self, path):
        self.path = path[0]

    def reload_work_book(self):
        self.wb = load_workbook(self.path)
        self.ws = self.wb.active
