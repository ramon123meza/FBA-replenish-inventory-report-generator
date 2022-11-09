'''

Created on May 16, 2022

@author: Arturo Cardona

'''

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import datetime
import pandas as pd
import numpy as np
from datetime import date
import time
# initalise the tkinter GUI
root = tk.Tk()

root.geometry("500x500")  # set the root dimensions
# tells the root to not let the widgets inside it determine its size.
root.pack_propagate(False)
root.resizable(0, 0)  # makes the root window fixed in size.

# Frame for TreeView
frame1 = tk.LabelFrame(root, text="Excel Data")
frame1.place(height=250, width=500)

# Frame for open file dialog
file_frame = tk.LabelFrame(root, text="Open File")
file_frame.place(height=200, width=500, rely=0.55, relx=0)
i = 0
# Buttons
button1 = tk.Button(file_frame, text="Browse", command=lambda: File_dialog())
button1.place(rely=0.85, relx=0)


button2 = tk.Button(file_frame, text="RestockReport",
                    command=lambda: Load_excel_data())
button2.place(rely=0.85, relx=0.30)

button4 = tk.Button(file_frame, text="PrintListItems",
                    command=lambda: Load_excel_data2())
button4.place(rely=0.85, relx=0.60)

# The file/file path text


label_file = ttk.Label(file_frame, text="No File Selected")

label_file.place(rely=0, relx=0)

# Treeview Widget
tv1 = ttk.Treeview(frame1)
# set the height and width of the widget to 100% of its container (frame1).
tv1.place(relheight=1, relwidth=1)

# command means update the yaxis view of the widget
treescrolly = tk.Scrollbar(frame1, orient="vertical", command=tv1.yview)
# command means update the xaxis view of the widget
treescrollx = tk.Scrollbar(frame1, orient="horizontal", command=tv1.xview)
# assign the scrollbars to the Treeview Widget
tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)
# make the scrollbar fill the x axis of the Treeview widget
treescrollx.pack(side="bottom", fill="x")
# make the scrollbar fill the y axis of the Treeview widget
treescrolly.pack(side="right", fill="y")

filelist = []


def File_dialog():
    """This Function will open the file explorer and assign the chosen file path to label_file"""
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"), ("All Files", "*.*")))

    if label_file["text"] == "No File Selected":

        label_file["text"] = filename
        filelist.append(filename)
    else:
        name = label_file["text"]

        label_file["text"] = name + "--" + filename
        filelist.append(filename)
    print(filelist)
    return None


def task1():
    da = datetime.datetime.now().date()

    res = (pd.Timestamp(da) + pd.DateOffset(days=7)).strftime('%Y-%m-%d')

    file = label_file["text"]
    file = file.split('--')
    print(file[0])
    Rr = pd.read_excel(file[0], parse_dates=['Recommended ship date'])
    Pl = pd.read_excel(file[1])

    res = res+" 00:00:00"
    checklist = ['LAN-RT-AP-BNC', 'FLG-N-CHE22', 'FLG-N-KH22', 'CH-N-NAS22']
    search = Pl['SKU']
    print(len(Rr))
    for i in range(0, len(Rr)):
        if (Rr['Merchant SKU'][i] in set(search) or Rr['Recommended ship date'][i] > res or Rr['Recommended replenishment qty'][i] <= 0):
            Rr.drop(i, inplace=True)
        elif Rr['Recommended replenishment qty'][i] > 0 and Rr['Total Units'][i] >= 12 and Rr['Merchant SKU'][i] not in checklist:
            Rr.drop(i, inplace=True)

    Rr = Rr.drop(['Country/Region Code', 'FNSKU',  'ASIN',
                 'Condition', 'Supplier', 'Supplier part no.', 'Currency code', 'Price',
                  'Sales last 30 days', 'Units Sold Last 30 Days',
                  'Inbound', 'Available', 'FC transfer', 'FC Processing',
                  'Customer Order', 'Unfulfillable', 'Working', 'Shipped', 'Receiving',
                  'Fulfilled by',
                  'Total days of supply (including units from open shipments)',
                  'Days of supply at Amazon fulfillment centers', 'Alert',
                  'Recommended action'], axis=1)
    return Rr


def task2():
    file = label_file["text"]
    Sr = pd.read_excel(file)
    for i in range(0, len(Sr)):
        name = Sr['fulfillment'][i]

        if name != "Seller" or Sr['sku'][i] == " " or 'Custom' in Sr['description'][i] or 'Customizable' in Sr['description'][i]:
            Sr.drop(i, inplace=True)
    Sr = Sr.drop(['settlement id', 'type', 'marketplace', 'account type', 'fulfillment', 'order city',
                  'order state', 'order postal', 'tax collection model', 'product sales',
                  'product sales tax', 'shipping credits', 'shipping credits tax',
                  'gift wrap credits', 'giftwrap credits tax', 'Regulatory Fee',
                  'Tax On Regulatory Fee', 'promotional rebates', 'date/time', 'description',
                  'promotional rebates tax', 'marketplace withheld tax', 'selling fees',
                  'fba fees', 'other transaction fees', 'other', 'total'], axis=1)
    Sr = Sr.groupby(['sku'])['quantity'].count()
    Sr = Sr[Sr > 5]
    Sr = Sr.sort_values(ascending=False)

    return Sr


def Load_excel_data2():
    """If the file selected is valid this will load the file into the Treeview"""
    df = task2()

    df.to_excel("FBA.xlsx")
    clear_data()
    load = pd.read_excel("FBA.xlsx")
    tv1["column"] = list(load.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        # let the column heading = column name
        tv1.heading(column, text=column)

    df_rows = df.to_numpy().tolist()  #turns the dataframe into a list of lists
    for row in df_rows:
        # inserts each list into the treeview. For parameters see https://docs.python.org/3/library/tkinter.ttk.html#tkinter.ttk.Treeview.insert
        tv1.insert("", "end", values=row)
    return None


def Load_excel_data():
    """If the file selected is valid this will load the file into the Treeview """
    df = task1()

    #time_1=time.strftime("%H%M%S")
    
    date_time = time.strftime("%m-%d-%Y %H-%M-%S")


    df.to_excel("FBAreplenish"+"("+(date_time)+")"+'.xlsx')
    clear_data()
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        # let the column heading = column name
        tv1.heading(column, text=column)

    df_rows = df.to_numpy().tolist()  # turns the dataframe into a list of lists
    for row in df_rows:
        # inserts each list into the treeview. For parameters see https://docs.python.org/3/library/tkinter.ttk.html#tkinter.ttk.Treeview.insert
        tv1.insert("", "end", values=row)
    return None


def clear_data():
    tv1.delete(*tv1.get_children())
    return None


root.mainloop()
