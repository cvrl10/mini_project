import tkinter as tk
from tkinter import ttk, filedialog, Spinbox
from pathlib import Path
import win32com.client as win32
from template import Template, Report
from tooltip import Tooltip
import pywintypes

DEFAULT_TEXT = 'Please select the folder containing your data file(s).'


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()

        self.root.title('PARTICLE SIZE DISTRIBUTION (PSD)')
        self.root.iconbitmap('img/blank.ico')

        self.root.geometry('500x110')
        self.root.resizable(False, False)

        self.root.columnconfigure(0, weight=0)
        self.root.columnconfigure(1, weight=1)

        self.root.rowconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)
        self.root.rowconfigure(2, weight=1)
        self.root.rowconfigure(3, weight=1)

        button = ttk.Button(self.root, text='DATA FOLDER', command=self._askdirectory, takefocus=False)
        button.grid(column=0, row=0, sticky='ew', padx=(10, 0))

        self.label = ttk.Label(self.root, text=DEFAULT_TEXT)
        self.label.grid(column=1, row=0, sticky='ew', padx=(0, 10))

        tooltip = Tooltip(self.label, DEFAULT_TEXT)
        tooltip.bind()

        button = ttk.Button(self.root, text='CREATE WORKBOOK', command=self._create_workbook, takefocus=False, width=19)
        button.grid(row=1, column=0, columnspan=2)

        self.STEM_EDS_ROW = tk.IntVar(value=0)

        label = ttk.Label(self.root, text='STEM-EDS')
        label.grid(row=2, column=0, columnspan=2, padx=(0, 65))
        spinbox = Spinbox(self.root, from_=0, to=118, width=3, textvariable=self.STEM_EDS_ROW)
        spinbox.grid(row=2, column=0, columnspan=2, padx=(90, 0))

        self.report_button = ttk.Button(self.root, text='CREATE REPORT', command=self._create_report, takefocus=False, width=19)
        self.report_button.grid(row=3, column=0, columnspan=2)
        self.report_button.state(['disabled'])

        self.root.bind('<Escape>', lambda _: self.root.destroy())

    def _askdirectory(self):
        data_folder = filedialog.askdirectory(parent=self.root, title='Select Folder')
        if data_folder:
            self.label.configure(text=f'{data_folder}')
            if self._has_workbook(data_folder):
                self.report_button.state(['!disabled'])
        else:
            self.label.configure(text=DEFAULT_TEXT)
            self.report_button.state(['disabled'])

    def _create_workbook(self):
        data_folder = self.label.cget('text')
        if data_folder == DEFAULT_TEXT:
            return

        data_folder = Path(data_folder)
        template = Template(data_folder=data_folder)
        template.create_sheets()
        self.report_button.state(['!disabled'])

    def _create_report(self):
        data_folder = self.label.cget('text')
        data_folder = Path(data_folder)
        file = data_folder/f'_{data_folder.name}.pptx'
        powerpoint = win32.Dispatch('PowerPoint.Application')

        try:
            presentation = powerpoint.Presentations(file.name)
            presentation.close()
        except pywintypes.com_error:
            #print(e)
            print('Exception was thrown')

        STEM_EDS = self.STEM_EDS_ROW.get()

        report = Report(data_folder=data_folder)
        report.create_slides(STEM_EDS=STEM_EDS)

    def _has_workbook(self, data_folder):
        file = Path(data_folder)
        file = file/f'_{file.name}.xlsm'
        return file.is_file()

    def run(self):
        self.root.deiconify()
        self.root.mainloop()
