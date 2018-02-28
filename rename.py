import os
from openpyxl import load_workbook
files=os.listdir()
workbook=load_workbook(r'..\摄影类.xlsx')
worksheet=workbook.active
zuopin=worksheet['c3:c86']
for file in files:
    for cell in zuopin:
        if cell[0].value in file:
            os.renames(file,str((zuopin.index(cell))+1)+file)
            break
        else:
            continue
##print(files)
