import xlrd
import readtable
import json
import collections
def readarray(book,sheet,begin,count,primary):
    cols=sheet.ncols
    extention=['array','table']
    ignore=['begin','count','']
    record_cols=[]
    record_type=[]
    record_name=[]
    result=[]
    global value
    try:
        for current_col in range(cols):
            type=sheet.cell_value(0,current_col)
            if not(type in ignore):
                record_cols.append(current_col)
                record_type.append(type)
                record_name.append(sheet.cell_value(1,current_col))
        nums=len(record_cols)
        if primary in record_name or primary==sheet.cell_value(1,0):
            isarray=0
        elif 'string' in record_type or 'int' in record_type or 'object' in record_type or 'float' in record_type:
            isarray=0
        else:
            isarray=1
        for row in range(begin-1,begin+count-1):
            item=collections.OrderedDict()
            for i in range(nums):
                if record_type[i] in extention:
                    childsheet=book.sheet_by_name(sheet.cell_value(row,record_cols[i]))
                    begin = int(sheet.cell_value(row, record_cols[i] + 1))
                    count = int(sheet.cell_value(row, record_cols[i] + 2))
                    if record_type[i]=='array':
                        if isarray:
                            item=readarray(book,childsheet,begin,count,'empty')[0]
                        else:
                            item[record_name[i]]=readarray(book,childsheet,begin,count,record_cols[i])[0]
                    else:
                        item[record_name[i]]=readtable.readtable(book,childsheet,begin,count,record_name[i])
                else:
                    value=sheet.cell_value(row,record_cols[i])
                    if record_type[i]=='object':
                        value=json.loads(value)
                    elif record_type[i]=='int':
                        value=int(value)
                    item[record_name[i]]=value
            result.append(item)
        return (result,isarray)
    except ValueError:
        print value
        raise ValueError
    

