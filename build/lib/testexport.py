from xlrd import *
import json
import readtable
def jsonexport(filename,dir,exportdir):
    xlsfile=dir+filename
    book=open_workbook(xlsfile)
    output=book.sheet_by_name('output')
    file=open(exportdir+filename[:-4]+'.json','w')
    result=readtable.readtable(book,output,3,1,'empty')
    result=json.dumps(result)
    file.write(result)
    file.close()

