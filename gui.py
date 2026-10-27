import tkinter as tk
from tkinter import ttk, filedialog

class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry('500x250')
        self.root.columnconfigure(0, weight=0)
        self.root.columnconfigure(1, weight=1)

        self.root.rowconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        button = ttk.Button(self.root, text='DATA FOLDER', command=lambda: self._data_folder())
        button.grid(column=0, row=0, sticky='ew', padx=(10,0))

        self.label = ttk.Label(self.root, text='')
        self.label.grid(column=1, row=0, sticky='ew', padx=(0,10))

        button = ttk.Button(self.root, text='CREATE WORKBOOK', command=lambda: self._create_workbook())
        button.grid(row=1, column=0, columnspan=2)

    def _data_folder(self):
        data_folder = filedialog.askdirectory(parent=self.root, title='data folder')
        self.label.configure(text=f'{data_folder}')

    def _create_workbook(self):
        print('clicked')

    def run(self):
        self.root.mainloop()



app = App()
app.run()