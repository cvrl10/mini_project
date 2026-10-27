from zipfile import ZipFile

with ZipFile("macro.xlsm") as wb:
    with open("vbaProject.bin", "wb") as f:
        f.write(wb.read("xl/vbaProject.bin"))
