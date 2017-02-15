from my_remove_file import my_copy_files
import shutil

s_dir = r'C:\Dmitry\GitHub\SReport\sreport_project\StrengthReport\StrengthReport\bin\Debug'
d_dir = r'C:\Dmitry\ExportSReport'
rep_dir_list = ['0408', '212247', 'big', 'new']

# copy group files
my_copy_files(s_dir, d_dir, '*.dll')

# copy single files
single_files = ['StrengthReport.exe', 'data.xls']

for p in single_files:
    full_path = ''.join([s_dir, '\\', p])
    dest = ''.join([d_dir, '\\', p])
    print('copy ', dest)                  
    shutil.copy(full_path, dest)

    

print('Export files is completed')
