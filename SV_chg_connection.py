# coding: utf8

import os, fnmatch
from Tkinter import Frame, Label, Entry, Button, Tk, TOP, RIGHT, LEFT, X, YES
import zipfile, shutil
import codecs

fields = 'File', 'Search', 'Replace'

def fetch(entvals):
#    print entvals
#    print ents
    entItems = entvals.items()
    for entItem in entItems:
        field = entItem[0]
        text  = entItem[1].get()
        print('%s: "%s"' % (field, text))

def findReplace(entvals):
#    print ents
    #file = entvals.get("File").get()
    file = "C:\Users\mcollet\Documents\PyConnect\Liasse.xlsx"
    directory = file.replace('.xlsx','')
    find = entvals.get("Search").get()
    replace = entvals.get("Replace").get()
    filePattern = "*.bin"

    try:
        os.remove(directory + "_test.xlsx")
    except:
        print("Can't remove the file or does not exist")
        pass
    
    with zipfile.ZipFile(file, 'r') as zip_ref:
        zip_ref.extractall(directory)

    for path, dirs, files in os.walk(os.path.abspath(directory)):
        for filename in fnmatch.filter(files, filePattern):
#            print filename
            filepath = os.path.join(path, filename)
            print filepath  # Can be commented out --  used for confirmation

            with codecs.open(filepath) as f:
                s = f.read()
            
            try:
                s = s.decode('utf-16')
                s = s.replace(find, replace)

                with codecs.open(filepath, mode="w", encoding="utf-16-le") as f:
                    try:
                        f.write(s)
                    except UnicodeDecodeError:
                        print("UnicodeDecodeError:{:s}".format(f))
                        return
            except:
                continue
    
    shutil.make_archive(directory+'_test', 'zip', directory)
    os.rename(directory + "_test.zip", directory + "_test.xlsx")
    shutil.rmtree(directory)

def makeform(root, fields):
    entvals = {}
    for field in fields:
        row = Frame(root)
        lab = Label(row, width=17, text=field+": ", anchor='w')
        ent = Entry(row)
        row.pack(side=TOP, fill=X, padx=5, pady=5)
        lab.pack(side=LEFT)
        ent.pack(side=RIGHT, expand=YES, fill=X)
        entvals[field] = ent
#        print ent
    return entvals

if __name__ == '__main__':
    root = Tk()
    root.title("Recursive Search & Replace")
    ents = makeform(root, fields)
#    print ents
    root.bind('<Return>', (lambda event, e=ents: fetch(e)))
    b1 = Button(root, text='Show', command=(lambda e=ents: fetch(e)))
    b1.pack(side=LEFT, padx=5, pady=5)
    b2 = Button(root, text='Execute', command=(lambda e=ents: findReplace(e)))
    b2.pack(side=LEFT, padx=5, pady=5)
    b3 = Button(root, text='Quit', command=root.quit)
    b3.pack(side=LEFT, padx=5, pady=5)
    root.mainloop()