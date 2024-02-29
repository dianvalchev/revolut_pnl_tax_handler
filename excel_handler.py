import openpyxl as xl


class ExcelHandler:

    def __init__(self, xl_file):
        self.result_file = "ResultFile.xlsx"
        self.header_row = None
        self.xl_file = xl_file
        self.data = self.open_pnl_file()

    def open_pnl_file(self):
        # load the pnl file with openpyxl
        wb = xl.load_workbook(self.xl_file)
        ws = wb.active

        # store the header row in a list
        self.header_row = [x.value for x in ws[1]]

        # store the remaining rows in a list
        data = list(ws.iter_rows(min_row=2, values_only=True))

        # create a list of dictionaries that has header data as keys and remaining rows' data as values
        # data = [{string: integer for string, integer in zip(self.header_row, row_data)} for row_data in row_data]

        return data

    def create_result_file(self):
        # define the results file name

        # create the file and delete the default sheet
        wb = xl.Workbook()
        ws = wb.active
        wb.remove(ws)

        # TODO: the parse logic goes here...
        for item in self.data:
            sheet_name = str(item[1].year)
            if sheet_name not in wb.sheetnames:
                wb.create_sheet(sheet_name)
                wb[sheet_name].append(self.header_row)

            wb[sheet_name].append(item)

        wb._sheets.sort(key=lambda sheet: sheet.title)
        wb.save(self.result_file)

    def add_bgn_pnl_to_result_file(self, data):
        wb = xl.load_workbook(self.result_file)
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            pnl_total = 0.00

            # add the new header row items
            for item in ['Date Acquired (BGN/USD)', 'Date Sold (BGN/USD)', 'Cost basis (BGN)', 'Amount (BGN)',
                         'Realised PnL (BGN)']:
                ws.cell(row=1, column=ws.max_column + 1).value = item

            # calculate and fill the PnL data in BGN
            for row in range(2, ws.max_row + 1):
                date_acquired = ws.cell(row=row, column=1).value.strftime('%Y-%m-%d')
                date_sold = ws.cell(row=row, column=2).value.strftime('%Y-%m-%d')

                da_fx_rate = ws.cell(row=row, column=9).value = data[date_acquired]
                ds_fx_rate = ws.cell(row=row, column=10).value = data[date_sold]

                cost_basis_usd = ws.cell(row=row, column=5).value
                amount_usd = ws.cell(row=row, column=6).value

                cost_basis_bgn = cost_basis_usd * da_fx_rate
                amount_bgn = amount_usd * ds_fx_rate

                ws.cell(row=row, column=11).value = cost_basis_bgn
                ws.cell(row=row, column=12).value = amount_bgn

                ws.cell(row=row, column=13).value = amount_bgn - cost_basis_bgn
                pnl_total += ws.cell(row=row, column=13).value

                ws.cell(row=row, column=5).number_format = "$ #,##0.00"
                ws.cell(row=row, column=6).number_format = "$ #,##0.00"
                ws.cell(row=row, column=7).number_format = "$ #,##0.00"
                ws.cell(row=row, column=9).number_format = "#,##0.00 лв."
                ws.cell(row=row, column=10).number_format = "#,##0.00 лв."
                ws.cell(row=row, column=11).number_format = "#,##0.00 лв."
                ws.cell(row=row, column=12).number_format = "#,##0.00 лв."
                ws.cell(row=row, column=13).number_format = "#,##0.00 лв."

            # write the total at the end of the sheet (=SUM)
            ws.cell(row=ws.max_row + 1, column=13).value = pnl_total
            ws.cell(row=ws.max_row, column=13).number_format = "#,##0.00 лв."
        wb.save(self.result_file)


if __name__ == "__main__":
    data = {'2019-11-27': 1.77657, '2020-01-28': 1.77722, '2019-11-25': 1.77674, '2021-09-14': 1.65552,
            '2022-06-10': 1.84896, '2021-11-15': 1.70904, '2020-02-28': 1.78175, '2020-06-05': 1.72624,
            '2021-08-16': 1.66143, '2022-02-07': 1.7086, '2020-01-29': 1.77787, '2019-12-06': 1.76296,
            '2021-01-22': 1.60868, '2020-07-02': 1.73297, '2020-07-06': 1.727, '2021-07-09': 1.64938,
            '2020-08-04': 1.66241, '2020-04-14': 1.78403, '2019-12-12': 1.75616, '2020-03-12': 1.74006,
            '2020-01-17': 1.76074, '2023-10-16': 1.85598, '2019-11-26': 1.7748, '2023-03-02': 1.84425,
            '2020-01-03': 1.75458, '2020-04-06': 1.81246, '2020-10-08': 1.66241, '2021-10-20': 1.68272,
            '2019-12-02': 1.77432, '2022-09-07': 1.97858, '2023-09-05': 1.8226, '2021-05-25': 1.59477,
            '2020-11-02': 1.67854, '2021-02-16': 1.61066, '2021-03-29': 1.65973, '2022-03-03': 1.76106,
            '2020-05-01': 1.7983}

    statement_file_name = "trading-pnl-statement_2019-11-01_2024-01-01_en-us_d59d3d.xlsx"  # 2019-2023 PnL
    my_handler = ExcelHandler(statement_file_name)
    my_handler.create_result_file()
    my_handler.add_bgn_pnl_to_result_file(data)
