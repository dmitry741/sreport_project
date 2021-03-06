from my_remove_file import my_copy_files
from my_remove_file import my_file_remove
import shutil
import os
from datetime import datetime

''' script deployes sreport project '''

start_time = datetime.now()

s_dir = r'C:\Dmitry\GitHub\sreport_project\StrengthReport\StrengthReport\bin\Debug'
d_dir = r'X:\DEN\RUDH\Technical\Technical_All\Pavlov\Расчет на прочность\ExportSReport'
rep_dir_list = ['0408', '212247', 'big', 'new']

# remove files from the current directory 
my_file_remove(d_dir, '*.*')

# remove or create files from (in) reports
for d in rep_dir_list:
    dd = ''.join([d_dir, '\\reports\\', d])    

    if os.path.exists(dd):
        print('remove all files from ', dd)
        my_file_remove(dd, '*.*')
    else:
        print('create ', dd)
        os.makedirs(dd)       

# copy group files to the application start
my_copy_files(s_dir, d_dir, '*.dll')

# copy single files to the application start
single_files = ['StrengthReport.exe', 'data.xls']

for p in single_files:
    full_path = ''.join([s_dir, '\\', p])
    dest = ''.join([d_dir, '\\', p])
    print('copy ', dest)                  
    shutil.copy(full_path, dest)

# copy reports
for d in rep_dir_list:
    dd = ''.join([d_dir, '\\reports\\', d])
    ss = ''.join([s_dir, '\\reports\\', d])
    print('copy reports from {} to {}'.format(ss, dd))
    my_copy_files(ss, dd, '*.mrt')


print('Export files is completed at: ', datetime.strftime(datetime.now(), "%Y.%m.%d %H:%M:%S"))
end_time = datetime.now()
print('Elapsed time: ', end_time - start_time)
