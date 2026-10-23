import pandas as pd
import re
import glob
from pathlib import Path
import xlsxwriter

path = r'100048497'
#path = r'100040000'
directory = Path(path)

print(directory)

class Template:
    def __init__(self, directory: Path):
        self.directory = directory
        self.writer = pd.ExcelWriter(self.directory / 'output.xlsx', engine='xlsxwriter')

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
        df = pd.read_excel(path, header=None, names=['Particle', 'Diameter (nm)'])

        df.to_excel(self.writer, sheet_name=path.stem, index=False)
        print(df)
        end_row = len(df)+1
        print(f'end_row = {end_row}')

        workbook = self.writer.book
        worksheet = self.writer.sheets[path.stem]

        worksheet.write('E1', 'min')
        worksheet.write('E2', df['Diameter (nm)'].min())

        worksheet.write('F1', 'max')
        worksheet.write('F2', df['Diameter (nm)'].max())

        worksheet.write('G1', 'avg')
        worksheet.write('G2', df['Diameter (nm)'].mean())

        worksheet.write('H1', 'std')
        worksheet.write('H2', df['Diameter (nm)'].std())

        worksheet.write('J1', 'bin_size')
        worksheet.write('J2', 2)

        worksheet.write('K1', 'Bin')
        worksheet.write_formula('K2', '=SEQUENCE(INT((P2-O2)/J2)+1,1,O2,J2)')
        worksheet.write('L1', 'Mid')
        worksheet.write_dynamic_array_formula('L2', '=(_xlfn.ANCHORARRAY(Q2)+_xlfn.ANCHORARRAY(K2))/2')

        worksheet.write('M1', 'Frequency')
        #worksheet.write_formula()

        worksheet.write('O1', 'start_bin')
        worksheet.write_formula('O2', '=FLOOR(E2,J2)')

        worksheet.write('P1', 'end_bin')
        worksheet.write_formula('P2', '=CEILING(F2,J2)')

        worksheet.write('Q1', 'data_for_mid_calc')
        worksheet.write_formula('Q2', '=SEQUENCE(INT((P2-O2)/J2)+1,1,O2+J2,J2)')

        #worksheet.write_dynamic_array_formula('M2', f'''=COUNTIFS(B2:B{end_row}, ">="&_xlfn.ANCHORARRAY(K2), B2:B{end_row}, "<"&(_xlfn.ANCHORARRAY(K2)+$J$2))''')
        worksheet.write_dynamic_array_formula('M2', f'''=COUNTIFS(B2:B{end_row}, ">="&(_xlfn.ANCHORARRAY(K2)-$J$2), B2:B{end_row}, "<"&_xlfn.ANCHORARRAY(K2))''')#used this equation to match exactly what Gyan does





    def create_sheets(self):
        for file in self.files:
            self._create_sheet(file)
        self.writer.close()

template = Template(directory)

print(template.files)

template.create_sheets()