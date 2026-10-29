import tkinter as tk
from tkinter import ttk, filedialog
from template import Template, Path
from tooltip import Tooltip

DEFAULT_TEXT = 'Please select the folder containing your data file(s).'

class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()

        self.root.title('PARTICLE SIZE DISTRIBUTION (PSD)')
        #self.root.iconbitmap('img/logo.ico')
        self.root.iconbitmap('img/blank.ico')

        self.root.geometry('500x70')
        self.root.resizable(False, False)

        self.root.columnconfigure(0, weight=0)
        self.root.columnconfigure(1, weight=1)

        self.root.rowconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        button = ttk.Button(self.root, text='DATA FOLDER', command=self._askdirectory, takefocus=False)
        button.grid(column=0, row=0, sticky='ew', padx=(10, 0))

        self.label = ttk.Label(self.root, text=DEFAULT_TEXT)
        self.label.grid(column=1, row=0, sticky='ew', padx=(0, 10))

        Tooltip(self.label, DEFAULT_TEXT)

        button = ttk.Button(self.root, text='CREATE WORKBOOK', command=self._create_workbook, takefocus=False)
        button.grid(row=1, column=0, columnspan=2)

    def _askdirectory(self):
        data_folder = filedialog.askdirectory(parent=self.root, title='Select Folder')
        if data_folder:
            self.label.configure(text=f'{data_folder}')
        else:
            self.label.configure(text=DEFAULT_TEXT)

    def _create_workbook(self):
        data_folder = self.label.cget('text')
        if data_folder == DEFAULT_TEXT:
            return
        data_folder = Path(data_folder)
        template = Template(data_folder=data_folder)
        template.create_sheets()

    def run(self):
        self.root.deiconify()
        self.root.mainloop()
