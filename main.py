import multiprocessing
import traceback
import xlsxwriter
import csv
import time
import math
import os
import pandas as pd
from tkinter import *
from tkinter import filedialog

RECORDS_PER_TAB = 1000000


def convert_csv_xlsx(file_path):
    try:
        file_path = os.path.join(os.getcwd(), file_path)
        print('Reading file: ' + file_path)
        f = open(file_path, encoding="utf_8")
        qty_rows = get_qty_rows_from_file(f)
        data = csv.DictReader(f)
        new_file_path = file_path.replace('.csv', '')
        new_file_path += '.xlsx'
        print('Writting new file')
        create_xlsx(data, new_file_path, qty_rows)
    except Exception as e:
        print('Error while reading file {} : {}'.format(file_path, e))


def create_xlsx(data, file_path, qty_records):
    try:
        current = 1
        tabs_number = math.ceil(qty_records / RECORDS_PER_TAB) if qty_records > RECORDS_PER_TAB else 1
        workbook = xlsxwriter.Workbook(file_path, {'constant_memory':True})
        print('Splitting {} records into {} tabs'.format(qty_records, tabs_number))
        tabs = []
        index_cell_format = workbook.add_format({
            'bold': True, 
            'bg_color': '#f5f5f5',
            'align': 'center', 
            'valign': 'vcenter'})
        headers_format = workbook.add_format({
            'italic': True,
            'bold': True, 
            'bg_color': '#b39e66',
            'align': 'center', 
            'valign': 'vcenter'})
        for index in range(tabs_number):
            print('{} / {}'.format(index + 1, tabs_number))
            worksheet = workbook.add_worksheet('Sheet{}'.format(index + 1))
            limit = RECORDS_PER_TAB * (index + 1) if (index < (tabs_number - 1)) else qty_records
            print('From {} To {}'.format(current, limit))
            worksheet.write_row(0, 1, data.fieldnames, headers_format)
            row_pos = 1
            for row_data in data:
                worksheet.write(row_pos, 0, current, index_cell_format)
                worksheet.write_row(row_pos, 1, row_data.values())
                row_pos += 1
                current += 1
                if current > limit:
                    break
        for tab in tabs:
            tab.start()
        for tab in tabs:
            tab.join()        
        print('Saving File')
        workbook.close()
        print('File created: {}'.format(file_path))
    except Exception as e:
        print('Error while creating xlsx file {} : {}'.format(file_path, e))


def get_qty_rows_from_file(f):
    qty_rows = 0
    for _ in f:
        qty_rows += 1
    f.seek(0)
    return qty_rows -1 if qty_rows > 0 else qty_rows

class App:
    def __init__(self):
        self.window = Tk()
        self.window.title("csv to xlsx converter")  # to define the title
        self.window.resizable(0, 0)
        self.window.configure(background='#ffffff')
        self.filename = None
        self.loaded = False
        self.label = Label(self.window , text = "File Explorer",
                                        width = 75, height = 4, background='#b39e66')
        btn_design = {'background':'#000000', 'foreground':'#ffffff', 'activebackground':'#b39e66',
                        'width':20}
        self.button_browse = Button(self.window, text="Browse File", command=self.browse_file, **btn_design)
        self.button_exit = Button(self.window, text="Exit", command=self.window.destroy, **btn_design)
        
        self.label.grid(row=0, column=0, pady=10, columnspan=3)
        self.button_browse.grid(row=1, column=0, columnspan=3)  
        self.button_exit.grid(row=2, column=1, pady=5)
            
    def run(self):
        self.window.mainloop()
        
    def browse_file(self):
        self.loaded = False
        self.filename = filedialog.askopenfilename(
                                        parent = self.window,
                                        initialdir = "/",
                                        defaultextension='.csv',
                                        filetypes=[('CSV file','*.csv'), ('All files','*.*')],
                                        title = "Select a File")
      
        if self.filename:
            try:
                self.label.configure(text = 'Processing file: {self.filename}')
                convert_csv_xlsx(self.filename)
                self.label.configure(text = 'File converted successfully \n\n File saved in container folder')
            except Exception as e:
                self.label.configure(text = 'Error: {}'.format(e))
        else:
            self.label.configure(text = 'No file selected')


if __name__ == '__main__':
    App().run()
