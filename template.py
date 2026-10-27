import pandas as pd
import re
from pathlib import Path
import os
from math import ceil, floor

path = r'100048497'
path = r'100040000'
directory = Path(path)

print(directory)

class Template:
    def __init__(self, data_folder: Path):
        self.directory = data_folder
        self.writer = pd.ExcelWriter(self.directory / f'{self.directory.name}_.xlsm', engine='xlsxwriter')
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
        print(f'inside _create_sheet {path}')
        sheet_name = path.stem
        df = pd.read_excel(path, header=None, names=['Particle (N)', 'Particle Size (nm)'])

        df.to_excel(self.writer, sheet_name=sheet_name, index=False)
        print(df)
        end_row = len(df)+1
        print(f'end_row = {end_row}')

        workbook = self.writer.book
        header_format = workbook.add_format({'bottom': 1,
                                             'bold': True})

        hide_format = workbook.add_format({'font_color': 'white'})

        worksheet = self.writer.sheets[sheet_name]
        worksheet.write(0, 0, 'Particle (N)', header_format)
        worksheet.write(0, 1, 'Particle Size (nm)', header_format)
        worksheet.autofit()

        worksheet.write('E1', 'min')
        worksheet.write('E2', df['Particle Size (nm)'].min())

        worksheet.write('F1', 'max')
        worksheet.write('F2', df['Particle Size (nm)'].max())

        worksheet.write('G1', 'avg')
        worksheet.write('G2', df['Particle Size (nm)'].mean())

        worksheet.write('H1', 'std')
        worksheet.write('H2', df['Particle Size (nm)'].std())

        worksheet.write('J1', 'bin_size')
        worksheet.write('J2', 2)

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


        x_axis = range(start_bin, end_bin+2, 2)
        print(len(x_axis))
        x_axis = list(x_axis)
        print(f'{sheet_name} x-axis: {x_axis}')

        print(f'{start_bin}')
        print(f'{end_bin}')

        #width = len(x_axis)

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
                    'color': '#ED7D31',
                    #'width': width
                }
            }
        })

        worksheet.insert_textbox('D5', f'''Average Particle Size: {series.mean():.2f} {chr(177)} {series.std():.2f} nm''',
                                 {'width': 400,
                                  'font': {'bold': True},
                                  'line': {'none': True},
                                  'fill': {'none': True}
            })
        #delete this
        '''
        chart.set_plotarea({
            'border': {
                'color': 'black',
                'width': 1
            }
        })'''
        worksheet.insert_chart('D5', chart)

    def create_sheets(self):
        for file in self.files:
            self._create_sheet(file)

        self.writer.close()
        os.startfile(self.directory/f'{self.directory.name}_.xlsm')


template = Template(data_folder=directory)
template.create_sheets()
