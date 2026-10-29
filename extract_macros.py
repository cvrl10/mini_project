from zipfile import ZipFile

with ZipFile('macros.xlsm') as wb:
    with open('vbaProject.bin', 'wb') as f:
        f.write(wb.read('xl/vbaProject.bin'))
