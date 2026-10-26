import pandas as pd
import re
import glob
from pathlib import Path
import os
from math import ceil, floor

path = r'100048497'
path = r'100040000'
directory = Path(path)

print(directory)

class Template:
    def __init__(self, directory: Path):
        self.directory = directory
        self.writer = pd.ExcelWriter(self.directory / 'output.xlsm', engine='xlsxwriter')
        self.writer.book.add_vba_project('vbaProject.bin')
    @property
    def files(self):
        source_files = []
        for file in glob.glob(str(self.directory.resolve()/'*')):
            if re.fullmatch(f'[0-9]+', Path(file).stem):
                source_files.append(Path(file).name)
        source_files.sort()
        return source_files

    def _create_sheet(self, file: str):
        path = self.directory / file
        print(f'inside _create_sheet {path}')
        df = pd.read_excel(path, header=None, names=['Particle (N)', 'Particle Size (nm)'])

        df.to_excel(self.writer, sheet_name=path.stem, index=False)
        print(df)
        end_row = len(df)+1
        print(f'end_row = {end_row}')

        start_bin = floor(df['Particle Size (nm)'].min()/2)*2
        end_bin = ceil(df['Particle Size (nm)'].max()/2)*2

        #x_axis =

        workbook = self.writer.book
        header_format = workbook.add_format({'bottom': 1,
                                             'bold': True})

        hide_format = workbook.add_format({'font_color': 'black'})

        worksheet = self.writer.sheets[path.stem]
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
        #worksheet.write_dynamic_array_formula('M2', f'''=COUNTIFS(B2:B{end_row}, ">="&_xlfn.ANCHORARRAY(K2), B2:B{end_row}, "<"&(_xlfn.ANCHORARRAY(K2)+$J$2))''')
        worksheet.write_dynamic_array_formula('M2', f'''=COUNTIFS(B2:B{end_row}, ">="&(_xlfn.ANCHORARRAY(K2)-$J$2), B2:B{end_row}, "<"&_xlfn.ANCHORARRAY(K2))''')#used this equation to match exactly what Gyan does

        worksheet.write('O1', 'start_bin', hide_format)
        worksheet.write_formula('O2', '=FLOOR(E2,J2)', hide_format)

        worksheet.write('P1', 'end_bin', hide_format)
        worksheet.write_formula('P2', '=CEILING(F2,J2)', hide_format)

        worksheet.conditional_format(f'Q1:Q{end_row}', {'type': 'no_blanks',
                                                        'format': hide_format}
        )#this to hide the data_for_mid_calc column

        worksheet.write('Q1', 'data_for_mid_calc')
        worksheet.write_dynamic_array_formula('Q2', '=SEQUENCE(INT((P2-O2)/J2)+1,1,O2+J2,J2)')

        worksheet.write('R1', 'zeros')
        worksheet.write_dynamic_array_formula('R2', '=SEQUENCE(INT((P2-O2)/J2)+1,1,0,0)')

        self._create_chart(sheet_name=path.stem, df=df)
        #self._create_chart_og(sheet_name=path.stem)

    def _create_chart(self, sheet_name, df):
        workbook = self.writer.book
        worksheet = self.writer.sheets[sheet_name]

        start_bin = floor(df['Particle Size (nm)'].min()/2) * 2
        end_bin = (ceil(df['Particle Size (nm)'].max()/2) * 2) + 2


        x_axis = range(start_bin, end_bin+2, 2)
        print(len(x_axis))
        x_axis = list(x_axis)
        print(f'{sheet_name} x-axis: {x_axis}')

        print(f'{start_bin}')
        print(f'{end_bin}')

        width = len(x_axis)
        print(f'{sheet_name} width: {width}')

        workbook.define_name(f'bin_{sheet_name}', f'={sheet_name}!$K$2:INDEX({sheet_name}!$K:$K,2+COUNT(_xlfn.ANCHORARRAY({sheet_name}!$K$2))-1)')
        workbook.define_name(f'mid_{sheet_name}', f'={sheet_name}!$L$2:INDEX({sheet_name}!$L:$L,2+COUNT(_xlfn.ANCHORARRAY({sheet_name}!$L$2))-1)')
        workbook.define_name(f'zeros_{sheet_name}', f'={sheet_name}!$R$2:INDEX({sheet_name}!$R:$R,2+COUNT(_xlfn.ANCHORARRAY({sheet_name}!$R$2))-1)')
        workbook.define_name(f'frequencies_{sheet_name}', f'={sheet_name}!$M$2:INDEX({sheet_name}!$M:$M,2+COUNT(_xlfn.ANCHORARRAY({sheet_name}!$M$2))-1)')


        chart = workbook.add_chart({
            'type': 'scatter',
            'subtype': 'smooth'
        })

        #column = workbook.add_chart(
            #{'type': 'column'}
        #)



        chart.set_legend({'none': True})
        chart.set_title({'name': f'PSD graph for {sheet_name}'})

        chart.set_x_axis({
            'name': 'Particle Size (nm)',
            'major_tick_mark': 'none',
            #'position_axis': 'between',
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
                'dash_type': 'solid'
            },
            'y_error_bars': {
                'type': 'percentage',
                'value': 100,
                'direction': 'minus',
                'end_style': 0,
                'line': {
                    'color': '#ED7D31',
                    'width': width
                }
            }
        })

        #column.add_series(
            #{'categories': f'={sheet_name}!bin_{sheet_name}',
                #'values': f'={sheet_name}!frequencies_{sheet_name}',
                #'gap': 0},
        #)

        #chart.combine(column)
        worksheet.insert_chart('D5', chart)

    def _create_chart_og(self, sheet_name):
        workbook = self.writer.book
        worksheet = self.writer.sheets[sheet_name]

        chart = workbook.add_chart(
            {'type': 'column'}
        )

        line = workbook.add_chart(
            {'type': 'line'}
        )

        workbook.define_name(f'bin_{sheet_name}', f'={sheet_name}!$K$2:INDEX({sheet_name}!$K:$K,2+COUNT(_xlfn.ANCHORARRAY({sheet_name}!$K$2))-1)')
        workbook.define_name(f'mid_{sheet_name}', f'={sheet_name}!$L$2:INDEX({sheet_name}!$L:$L,2+COUNT(_xlfn.ANCHORARRAY({sheet_name}!$L$2))-1)')
        workbook.define_name(f'zeros_{sheet_name}', f'={sheet_name}!$R$2:INDEX({sheet_name}!$R:$R,2+COUNT(_xlfn.ANCHORARRAY({sheet_name}!$R$2))-1)')
        workbook.define_name(f'frequencies_{sheet_name}', f'={sheet_name}!$M$2:INDEX({sheet_name}!$M:$M,2+COUNT(_xlfn.ANCHORARRAY({sheet_name}!$M$2))-1)')

        '''
        chart.add_series(
            {'categories': f'={sheet_name}!bin_{sheet_name}',
                'values': f'={sheet_name}!zeros_{sheet_name}',
                'gap': 6},
        )
        '''

        chart.add_series(
            {'categories': f'={sheet_name}!mid_{sheet_name}',
                'values': f'={sheet_name}!frequencies_{sheet_name}',
                'fill': {'color': '#ED7D31'},
                #'overlap': 100,
                'gap': 6},
        )

        line.add_series(
            {'categories': f'={sheet_name}!mid_{sheet_name}',
             'values': f'={sheet_name}!frequencies_{sheet_name}',
             'smooth': True,
             'line': {
                 'color': 'red'},
             'marker': {'type': 'none'}}
        )

        chart.set_legend({'none': True})
        chart.set_title({'name': f'PSD graph for {sheet_name}'})

        chart.set_x_axis({
            'name': 'Particle Size (nm)',
            'major_tick_mark': 'none',
            #'position_axis': 'on_tick',
            'position_axis': 'between',
            'min': 0,
        })
        chart.set_y_axis({
            'name': 'Frequency',
            'major_tick_mark': 'none',
            'min': 0,
            'major_gridlines': {'visible': False}
        })

        line.set_y_axis({
            'position_axis': 'on_tick'
        })

        chart.combine(line)
        #line.combine(chart)
        #worksheet.insert_chart('E10', line)
        worksheet.insert_chart('D5', chart)

    def create_sheets(self):
        for file in self.files:
            self._create_sheet(file)

        self.writer.close()
        os.startfile(self.directory/'output.xlsm')


template = Template(directory)
template.create_sheets()
