
from docxx import Document
import os 
import shutil
import argparse

pser = argparse.ArgumentParser()
pser.add_argument('input_file', help='ooxml document')
args = pser.parse_args()

def subparts(parts, level=0):
    for k, v in parts.related_parts.items():
        print("　"*level, k, v.partname)
        subparts(v,level+1)

#subparts(xlsx)

xlsx = Document(args.input_file)
for sheet in xlsx.workbook.sheets:
    print(" --- sheet '{}' --- ".format(sheet.name))
    for cell in sheet.get_cells(row=5):
        print("{:3}: text={}, v={}, type={}, st={}".format(cell.ref, cell.text, cell.value, cell._element.t, cell._element.s))

xlsx.save("test_style.xlsx")