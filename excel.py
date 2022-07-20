# +
from RPA.Excel.Files import Files


class WorkWithExcel:

    def __init__(self):
        self.excel = Files()

    def createFile(self, path):
        try:
            self.excel.create_workbook(path=path, fmt='xlsx')
            self.excel.save_workbook()
        finally:
            self.excel.close_workbook()

    def renameSheet(self, path, old_name, new_name):
        try:
            self.excel.open_workbook(path)
            self.excel.rename_worksheet(old_name, new_name)
            self.excel.save_workbook()
        finally:
            self.excel.close_workbook()

    def appendRow(self, data, path, sheet):
        try:
            self.excel.open_workbook(path)
            self.excel.append_rows_to_worksheet(data, sheet)
            self.excel.save_workbook()
        finally:
            self.excel.close_workbook()

    def createSheet(self, path, name):
        try:
            self.excel.open_workbook(path)
            self.excel.create_worksheet(name)
            self.excel.save_workbook()
        finally:
            self.excel.close_workbook()
