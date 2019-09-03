# Export a list of all files from a target path to an Excel Spreadsheet
# @exomachine 2019


import glob
import xlwt
import ntpath
from tempfile import TemporaryFile

path = 'U:\\'       #Input path
savePath = "Q:\\Folder Content.xls"     #Output path

rawData = []
book1 = xlwt.Workbook() 
sheet1 = book1.add_sheet('sheet1')

rawData.append('List of folders')
for d in glob.glob(path + "**/", recursive=False):
    rawData.append (d)
rawData.append('List of Folder and Files')
for f in glob.glob(path + "*", recursive=False):
    rawData.append (f)

#for f in rawData:
#    print(f)

def path_leaf(path): #trim path for filename or folder name
    head, tail = ntpath.split(path)
    return tail or ntpath.basename(head)

for i, e in enumerate(rawData):
    sheet1.write(i,0,path_leaf(e))
    sheet1.write(i,1,e)

book1.save(savePath)
book1.save(TemporaryFile())

print("done")
