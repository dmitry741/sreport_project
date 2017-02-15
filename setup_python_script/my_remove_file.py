''' module has methods for working with dir and files '''
import os
import shutil
import fnmatch

def my_copy_files (s_dir, d_dir, file_mask):
    ''' method copies files from s_dir to d_dir according to file_mask '''
    
    if not os.path.exists(d_dir):
        os.makedirs(d_dir)

    list_dir = os.listdir(s_dir)
    files = filter(lambda x: fnmatch.fnmatch(x, file_mask), list_dir)

    for p in files:
        ff = ''.join([s_dir, '\\', p])
        if os.path.isfile(ff):
            dd = ''.join([d_dir, '\\', p])
            print('copy ', dd)                        
            shutil.copy(ff, dd)

def my_remove_files(s_dir, file_mask):
    ''' Remove files from s_dir according to file_mask '''
    
    list_dir = os.listdir(s_dir)
    files = filter(lambda x: fnmatch.fnmatch(x, file_mask), list_dir)

    for p in files:
        ff = ''.join([s_dir, '\\', p])
        print('remove ', ff)                        

        result, reason = os.remove(ff)
  
        if result:
            print('File {} removed'.format(ff))
        else:
            print('Cannot remove file: ', reason)


