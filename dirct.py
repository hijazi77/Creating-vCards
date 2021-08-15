import csv
import sys
from openpyxl import Workbook
from openpyxl.reader.excel import load_workbook

name = [] 
phone = []
first = []
last = []

wb = load_workbook('auto.xlsx')
ws = wb.active


c = 1
for i in range(807):
  name.append(ws[f'A{c}'].value)
  first.append(ws[f'B{c}'].value)
  last.append(ws[f'C{c}'].value)
  phone.append(ws[f'D{c}'].value)
  c+=1
  
print(last[1]+" and "+phone[1])


file =28

for i in range(4):
  with open(f'folder/file{file}.vcf', 'w' , encoding="utf-8") as myfile:
    count = 0
    for contact in range(249):
      myfile.write( 'BEGIN:VCARD' + "\n")
      myfile.write( 'VERSION:2.1' + "\n")
      myfile.write( 'N:' + str(last[count]) + ';' + (first[count]) + "\n")
      myfile.write( 'FN:' +str(name[count])+ "\n")
      myfile.write( 'ORG:' + 'IBM' + "\n")
      myfile.write( 'TEL;CELL:' + str(phone[count])+ "\n")
      myfile.write( 'END:VCARD' + "\n")
      myfile.write( "\n")
      name.pop(count)
      first.pop(count)
      last.pop(count)
      phone.pop(count)
      count+= 1
    myfile.close()
  file +=1
  print(f'file{file} has been created')

    