from openpyxl import Workbook, load_workbook


class XlsxData:
    def __init__(self):
        self.wb = load_workbook('test.xlsx')
        self.ws = self.wb.active

        #self.ws['A2'].value = 'col add'
    def row_xlsx_add(self, row):
        self.ws.insert_rows(row)

    def col_xlsx_add(self, col):
        self.ws.insert_cols(col)

    def row_xlsx_del(self, row):
        self.ws.delete_rows(row)

    def col_xlsx_del(self, col):
        self.ws.delete_cols(col)

    def save_table(self):
        self.wb.save('test.xlsx')

    def get_working_sheet(self):
        return self.ws

    def reload_work_book(self):
        self.wb.remove_sheet(self.ws)
        self.wb = load_workbook('test.xlsx')
        self.ws = self.wb.active
