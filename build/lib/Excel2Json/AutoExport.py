import os
import testexport
import time
import xlrd
import sys
def listdir(dirname,pos_conditions,neg_conditions):
    result=[]
    list=os.listdir(dirname)
    for filename in list:
        check=1
        for condition in pos_conditions:
            if not condition in filename:
                check=0
                break
        for condition in neg_conditions:
            if condition in filename:
                check=0
                break
        if check==1:
            result.append(filename)
    return result
def jsonexport(dir='..\\config_common\\',expdir='..\\ramonexport\\',pos_conditions=['xls'],neg_conditions=['string','definition']):
    start=time.time()
    files=listdir(dir,pos_conditions,neg_conditions)
    nums =len(files)
    i=0
    for file in files:
        try:
            testexport.jsonexport(file,dir,expdir)
            i+=1
            print '%s  ( %d / %d )    '%(file,i,nums),time.time()-start,'s'
        except ValueError:
            print file,'went wrong, Type'
        except xlrd.biffh.XLRDError:
            print 'no output sheet in %s' %(file)
        except:
            print file,'went wrong'
    print 'Done',time.time()-start
if __name__ == '__main__':
    jsonexport()
