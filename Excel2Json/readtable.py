import xlrd
import readarray
import json
import collections
def isnumber(string):
    number=[str(i) for i in range(10)]
    number+=['.']
    for i in string:
        if not(i in number):
            return False
        else:
            return True
def interger(string):
    i=0
    for char in string:
        if char=='.':
            break
        i+=1
    return int(string[0:i])

def readtable(book,sheet,begin,count,primary):
    cols=sheet.ncols
    extention=['array','table']
    ignore=['begin','count','']
    record_cols=[]
    record_type=[]
    record_name=[]
    result=collections.OrderedDict()
    primary_col=0
    hasprimary=0
    global value
    if sheet.cell_value(0,1)==primary:
        hasprimary=1
    try:
        for current_col in range(cols):
            type=sheet.cell_value(0,current_col)
            if not(type in ignore):
                name=sheet.cell_value(1,current_col)
                if primary==name and hasprimary==0:
                    primary_col=current_col
                    hasprimary=1
                    if primary!='empty':
                        record_name.append(name)
                        record_cols.append(current_col)
                        record_type.append(type)
                else:
                    record_name.append(name)
                    record_cols.append(current_col)
                    record_type.append(type)
        nums=len(record_cols)
        for row in range(begin-1,begin+count-1):
            item=collections.OrderedDict()
            for i in range(nums):
                if record_type[i] in extention:
                    sheet_name=sheet.cell_value(row,record_cols[i])
                    childsheet=book.sheet_by_name(sheet_name)
                    begin = int(sheet.cell_value(row, record_cols[i] + 1))
                    count = int(sheet.cell_value(row, record_cols[i] + 2))
                    if record_type[i]=='array':
                        tmp=readarray.readarray(book,childsheet,begin,count,record_name[i])
                        if tmp[1]==1 and hasprimary==0 and count==1:
                            if not record_name[i] in item:
                                item[record_name[i]]=[]
                            item[record_name[i]].append(tmp[0])
                        else:
                            item[sheet_name]=tmp[0]
                    else:
                        item[sheet_name]=readtable(book,childsheet,begin,count,record_name[i])
                else:
                    value = sheet.cell_value(row, record_cols[i])
                    if record_type[i] == 'object':
                        value = json.loads(value)
                    elif record_type[i] == 'int':
                        value = int(value)
                    elif record_type[i]=="string":
                        try:
                            value=int(value)
                            value=str(value)
                        except ValueError:
                            value=str(value)
                    item[record_name[i]] = value
            if primary_col==-1:
                return item
            primary_name=str(sheet.cell_value(row,primary_col))
            if isnumber(primary_name):
                primary_name=interger(primary_name)
            if hasprimary:
                result[primary_name]=item
            else:
                result.update(item)
        return result
    except ValueError:
        print value
        raise ValueError
