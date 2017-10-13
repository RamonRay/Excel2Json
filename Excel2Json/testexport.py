# -*- coding: utf-8 -*-
from xlrd import *
import json
import readtable
import codecs
def singlejsonexport(filename,dir,expdir):
    xlsfile=dir+filename
    book=open_workbook(xlsfile)
    output=book.sheet_by_name('output')
    file=codecs.open(expdir+filename[:-4]+'.json','w','utf-8')
    result=readtable.readtable(book,output,3,1,'empty')
    result=json.dumps(result,ensure_ascii=False)
    result=result.encode('utf-8')
    file.write(result)
    file.close()

