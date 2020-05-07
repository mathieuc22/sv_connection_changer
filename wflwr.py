import os, shutil

zip_name = r'C:\Users\mcollet\Documents\PyConnect\Liasse_2'
directory_name = r'C:\Users\mcollet\Documents\PyConnect\Liasse'

shutil.make_archive(zip_name, 'zip', directory_name)

os.rename(zip_name + ".zip", zip_name + ".xlsx")