import os
from subprocess import *
import time
import sys
try:
    path=os.getcwd()
    os.system('git pull -v')
    os.system('git add config_common/\\*.xls')
    os.system('git commit -m "AutoDeploy"')
    os.system('git push')
    try:
        p = Popen('call export_all.bat', stdout=PIPE, shell=True)
        while 1:
            line=p.stdout.readline()
            time.sleep(0.01)
            if line!='':
                print line
            if 'MOKAK' in line and 'FINISH' in line:
                break
        p.terminate()
        raise Exception
    except:
        pass
    os.system('git pull -v')
    os.chdir(path+"\\tools")
    os.system('echo off')
    os.system('DownloadListMaker.exe')
    os.system('git add download_list.json')
    os.system('git commit -m "AutoMD5"')
    os.system('git push')
    print 'press Ctrl+C to Exit'
    raise Exception
except:
    sys.exit(0)





    
