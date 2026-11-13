import pandas as pd
import re
from pathlib import Path
import os
from math import ceil, floor
from pptx import Presentation
from pptx.dml.color import RGBColor
import win32com.client as win32
import pywintypes


class Report:
    def __init__(self, data_folder: Path, samples):
        self.directory = data_folder
        self.samples = samples
        self.presentation = Presentation('report_template.pptx')

    def extract_chart(self):
        directory = Path(self.directory/'chart(s)')
        directory.mkdir(exist_ok=True)

        workbook = self.directory/f'_{self.directory.name}.xlsm'
        excel = win32.Dispatch('Excel.Application')

        try:
            print(f'inside excel.Workbooks, workbook.name: {workbook.name}')
            wb = excel.Workbooks(workbook.name)
            excel.Run(f'{wb.Name}!Module3.ExportGroupedChartTextBox')
            print('workbook opened')
        except pywintypes.com_error:
            print('workbook not opened')
            wb = excel.Workbooks.Open(workbook.resolve().__str__())
            excel.Run(f'{wb.Name}!Module3.ExportGroupedChartTextBox')
            wb.Close(SaveChanges=False)

    def _create_background_slide(self):
        presentation = self.presentation
        slide = presentation.slides[1]

        table = slide.shapes[0].table
        cell = table.cell(2, 1)
        cell.text = ''
        paragraph = cell.text_frame.paragraphs[0]
        samples = paragraph.add_run()
        samples.text = ', '.join(Path(file).stem for file in self.samples)
        samples.font.color.rgb = RGBColor(255, 0, 0)

    def _create_chart_slide(self, sample: str):
        presentation = self.presentation
        layout = presentation.slides[3].slide_layout
        slide = presentation.slides.add_slide(layout)

        paragraph = slide.placeholders[0].text_frame.paragraphs[0]
        text = paragraph.add_run()
        text.text = 'Particle Size Distribution (PSD) Graph for: '

        text = paragraph.add_run()
        text.text = sample
        text.font.bold = True
        text.font.color.rgb = RGBColor(255, 0, 0)
        '''
        for ph in slide.placeholders:
            fmt = ph.placeholder_format
            print(fmt.idx, ph.name, fmt.type)'''
        image = slide.placeholders[21]#do not hardcode this, template is subject to change or you can name them
        image.insert_picture((self.directory/f'chart(s)/{sample}.png').resolve().__str__())

        #df = pd.read_excel(path, header=None, names=['Particle (N)', 'Particle Size (nm)'])
        paragraph = slide.placeholders[19].text_frame.paragraphs[0]
        text = paragraph.add_run()
        text.text = 'A total of 200 particles were analyzed to determine the average particle size of the catalyst'

    def _create_image_slide(self, sample: str):
        presentation = self.presentation
        layout = presentation.slides[2].slide_layout
        slide = presentation.slides.add_slide(layout)
        paragraph = slide.placeholders[0].text_frame.paragraphs[0]
        text = paragraph.add_run()
        text.text = 'STEM images for: '

        text = paragraph.add_run()
        text.text = sample
        text.font.bold = True
        text.font.color.rgb = RGBColor(255, 0, 0)

    def _create_slides(self):
        self._create_background_slide()
        self.extract_chart()#make this private
        for sample in [Path(file).stem for file in self.samples]:
            self._create_image_slide(sample)
            self._create_chart_slide(sample)

    def create(self):
        self._create_slides()
        self.presentation.save(self.directory/f'_{self.directory.name}.pptx')
        os.startfile(self.directory/f'_{self.directory.name}.pptx')


class Template:
    def __init__(self, data_folder: Path):
        self.directory = data_folder
        self.writer = pd.ExcelWriter(self.directory / f'_{self.directory.name}.xlsm', engine='xlsxwriter')
        self.writer.book.add_vba_project('vbaProject.bin')

    @property
    def files(self):
        source_files = []
        for path in self.directory.resolve().glob('*'):
            if re.fullmatch(f'[0-9]+', path.stem):
                source_files.append(path.name)
        source_files.sort()
        return source_files

    def _create_sheet(self, file: str):
        path = self.directory / file
        sheet_name = path.stem
        df = pd.read_excel(path, header=None, names=['Particle (N)', 'Particle Size (nm)'])
        df.to_excel(self.writer, sheet_name=sheet_name, index=False)

        end_row = len(df)+1

        workbook = self.writer.book
        header_format = workbook.add_format({'bottom': 1,
                                             'bold': True})

        hide_format = workbook.add_format({'font_color': 'white'})

        worksheet = self.writer.sheets[sheet_name]
        worksheet.write('A1', 'Particle (N)', header_format)
        worksheet.write('B1', 'Particle Size (nm)', header_format)
        worksheet.autofit()

        worksheet.write('E1', 'min')
        worksheet.write('E2', df['Particle Size (nm)'].min())

        worksheet.write('F1', 'max')
        worksheet.write('F2', df['Particle Size (nm)'].max())

        worksheet.write('G1', 'avg')
        worksheet.write('G2', df['Particle Size (nm)'].mean())

        worksheet.write('H1', 'std')
        worksheet.write('H2', df['Particle Size (nm)'].std())

        worksheet.write('J1', 'bin_size', workbook.add_format({
            'bg_color': '#000000',
            'font_color': '#FFFFFF',
            'italic': True,
            'bold': True
        }))

        worksheet.write('J2', 2, workbook.add_format({
            'italic': True,
            'bold': True
        }))

        worksheet.write('K1', 'Bin')
        worksheet.write_formula('K2', '=SEQUENCE(INT((P2-O2)/J2)+1,1,O2,J2)')

        worksheet.write('L1', 'Mid')
        worksheet.write_dynamic_array_formula('L2', '=(_xlfn.ANCHORARRAY(Q2)+_xlfn.ANCHORARRAY(K2))/2')

        worksheet.write('M1', 'Frequency')
        worksheet.write_dynamic_array_formula('M2', f'''=COUNTIFS(B2:B{end_row}, ">="&_xlfn.ANCHORARRAY(K2), B2:B{end_row}, "<"&(_xlfn.ANCHORARRAY(K2)+$J$2))''')
        #worksheet.write_dynamic_array_formula('M2', f'''=COUNTIFS(B2:B{end_row}, ">="&(_xlfn.ANCHORARRAY(K2)-$J$2), B2:B{end_row}, "<"&_xlfn.ANCHORARRAY(K2))''')#used this equation to match exactly what Gyan does

        worksheet.write('O1', 'start_bin', hide_format)
        worksheet.write_formula('O2', '=FLOOR(E2,J2)', hide_format)

        worksheet.write('P1', 'end_bin', hide_format)
        worksheet.write_formula('P2', '=CEILING(F2,J2)', hide_format)

        worksheet.conditional_format(f'Q1:Q{end_row}', {'type': 'no_blanks',
                                                        'format': hide_format}
        )#this to hide the data_for_mid_calc column

        worksheet.write('Q1', 'data_for_mid_calc')
        worksheet.write_dynamic_array_formula('Q2', '=SEQUENCE(INT((P2-O2)/J2)+1,1,O2+J2,J2)')

        self._create_chart(sheet_name=sheet_name, series=df['Particle Size (nm)'])

    def _create_chart(self, sheet_name: str, series: pd.Series):
        workbook = self.writer.book
        worksheet = self.writer.sheets[sheet_name]

        start_bin = floor(series.min()/2) * 2
        end_bin = (ceil(series.max()/2) * 2)

        workbook.define_name(f'mid_{sheet_name}', f'={sheet_name}!$L$2:INDEX({sheet_name}!$L:$L,2+COUNT(_xlfn.ANCHORARRAY({sheet_name}!$L$2))-1)')
        workbook.define_name(f'frequencies_{sheet_name}', f'={sheet_name}!$M$2:INDEX({sheet_name}!$M:$M,2+COUNT(_xlfn.ANCHORARRAY({sheet_name}!$M$2))-1)')

        chart = workbook.add_chart({
            'type': 'scatter',
            'subtype': 'smooth'
        })

        chart.set_legend({'none': True})
        chart.set_title({'name': f'PSD graph for {sheet_name}'})

        chart.set_x_axis({
            'name': 'Particle Size (nm)',
            'major_tick_mark': 'none',
            'min': start_bin,
            'max': end_bin,
            'major_unit': 2
        })

        chart.set_y_axis({
            'name': 'Frequency',
            'major_tick_mark': 'none',
            'min': 0,
            'major_gridlines': {'visible': False}
        })

        chart.add_series({
            'categories': f'={sheet_name}!mid_{sheet_name}',
            'values': f'={sheet_name}!frequencies_{sheet_name}',
            'line': {
                'color': 'red',
                'dash_type': 'solid',
                'width': 0.5
            },
            'y_error_bars': {
                'type': 'percentage',
                'value': 100,
                'direction': 'minus',
                'end_style': 0,
                'line': {
                    'color': '#ED7D31'
                }
            }
        })

        worksheet.insert_textbox('D5', f'''Average Particle Size: {series.mean():.2f} {chr(177)} {series.std():.2f} nm''',
                                 {'width': 300,
                                  'font': {'bold': True},
                                  'line': {'none': True},
                                  'fill': {'none': True}})

        chart.set_chartarea({'border': {'none': True}})#

        worksheet.insert_chart('D5', chart)

    def create_sheets(self):
        for file in self.files:
            self._create_sheet(file)

        self.writer.close()
        os.startfile(self.writer.book.filename.name)


path = '100040001'
path = Path(path)

#template = Template(path)
#template.writer.close()
#template.create_sheets()
#'''
samples = ['200123417', '200123418']
#report = Report(path, samples=template.files)
report = Report(path, samples=samples)
#report.extract_chart()
report.create()
#'''


