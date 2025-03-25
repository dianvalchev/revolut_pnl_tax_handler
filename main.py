from datetime import datetime
import run_spider
from run_spider import RunSpider
from excel_handler import ExcelHandler

import tkinter as tk
from tkinter import filedialog


def browse_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        entry_path.delete(0, tk.END)
        entry_path.insert(0, file_path)
        loading_label.config(text='Press "Generate"')

def press():
    loading_label.config(text="Working...")
    root.after(1,generate)

def generate():

    file_path = entry_path.get()
    statement_file_name = file_path

    my_handler = ExcelHandler(statement_file_name)

    # extract all the "acquired" and "sold" dates from the file and put them in a sorted list for parsing
    date_acquired_set = set()
    date_sold_set = set()

    dividends_reached = False
    for item in my_handler.data:
        if item[0] is None:
            break

        date_sold_set.add(datetime.strftime(item[1], "%Y-%m-%d"))
        date_acquired_set.add(datetime.strftime(item[0], "%Y-%m-%d"))

    #old way... not used anymore
    # date_acquired_set = set([datetime.strftime(date[0], "%Y-%m-%d") for date in my_handler.data])
    # date_sold_set = set([datetime.strftime(date[1], "%Y-%m-%d") for date in my_handler.data])

    date_list = sorted(list(date_sold_set.union(date_acquired_set)))

    # run the scrapy spider to retrieve all the FX rates for the corresponding dates
    my_spider = RunSpider(date_list)
    my_spider.run()
    # create a new Excel file with the transactions in BGN and PnL calculation
    if run_spider.result:
        my_handler.create_result_file()
        my_handler.add_bgn_pnl_to_result_file(run_spider.result)
        loading_label.config(text=" Done. ")
    else:
        print("run_spider result empty.")


# Create main window
root = tk.Tk()
root.title("File Generator")
root.geometry("300x200")

# Browse form
frame = tk.Frame(root)
frame.pack(pady=20)

label_path = tk.Label(frame, text="File Path:")
label_path.grid(row=0, column=0)

entry_path = tk.Entry(frame, width=20)
entry_path.grid(row=0, column=1)

button_browse = tk.Button(frame, text="Browse...", command=browse_file)
button_browse.grid(row=0, column=2)

# Generate button
button_generate = tk.Button(root, text="Generate", command=press)
button_generate.pack()

# Loading label
loading_label = tk.Label(root, bd=1, relief="solid", width=40, text="   Select Revolut PnL statement file...   ")
loading_label.pack(pady=(30,0))

empty_row = tk.Frame(height=50)
empty_row.pack()

root.mainloop()
