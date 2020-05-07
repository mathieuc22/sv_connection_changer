import zipfile
with zipfile.ZipFile('C:\Users\mcollet\Documents\PyConnect\Liasse budgétaire 2019.xlsx', 'r') as zip_ref:
    zip_ref.extractall('C:\Users\mcollet\Documents\PyConnect\')