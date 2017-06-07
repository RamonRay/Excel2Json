from Excel2Json import jsonexport
import os
dir='config_common\\'
if not os.path.exists(dir):
    print 'Dir doesn\'t exist'
expdir='checkexport\\'
if not os.path.exists(expdir):
    os.makedirs(expdir)
jsonexport(dir,expdir)
os.system('pause')
