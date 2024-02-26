from datetime import datetime
from excel_handler import ExcelHandler
from run_spider import RunSpider

# provide the PnL file from Revolut
statement_file_name = "trading-pnl-statement_2019-11-01_2024-01-01_en-us_d59d3d.xlsx"  # 2019-2023 PnL
my_handler = ExcelHandler(statement_file_name)

# extract all the "acquired" and "sold" dates from the file and put them in a sorted list
date_acquired_set = set([datetime.strftime(date["Date acquired"], "%Y-%m-%d") for date in my_handler.data])
date_sold_set = set([datetime.strftime(date["Date sold"], "%Y-%m-%d") for date in my_handler.data])
date_list = sorted(list(date_sold_set.union(date_acquired_set)))

# run the spider to retrieve all the FX rates for the corresponding dates
my_spider = RunSpider(date_list)
my_spider.run()
print("Final result in bnb_spider: ", my_spider.result)

# TODO: create the result file with the collected data:
