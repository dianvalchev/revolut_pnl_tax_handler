import openpyxl as xl


class ExcelHandler:

    def __init__(self, xl_file):
        self.xl_file = xl_file
        self.data = self.open_pnl_file()
        self.header_row = []

    def open_pnl_file(self):
        # load the pnl file with openpyxl
        wb = xl.load_workbook(self.xl_file)
        ws = wb.active

        # store the header row in a list
        self.header_row = [x.value for x in ws[1]]

        # store the remaining rows in a list
        row_data = list(ws.iter_rows(min_row=2, values_only=True))

        # create a list of dictionaries that has header data as keys and remaining rows' data as values
        data = [{string: integer for string, integer in zip(self.header_row, row_data)} for row_data in row_data]

        return data


if __name__ == "__main__":
    statement_file_name = "trading-pnl-statement_2019-11-01_2024-01-01_en-us_d59d3d.xlsx"  # 2019-2023 PnL
    my_handler = ExcelHandler(statement_file_name)
    print(my_handler.data)
