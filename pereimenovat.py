#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      oskin_rn
#
# Created:     26.05.2022
# Copyright:   (c) oskin_rn 2022
# Licence:     <your licence>
#-------------------------------------------------------------------------------
from xml.dom import minidom
import os
from tkinter import filedialog
from tkinter import messagebox
    

def ren(file, it, dirpath):
    newfile = it + '_' + file.split('_')[0].strip() + '.xls'
    os.rename(os.path.join(dirpath, file), os.path.join(dirpath, newfile))
    #print('Новый файл:', newfile)
    print(file)

def print_hi(name):
    directory  = filedialog.askdirectory()
    for dirpath, dirnames, filenames in os.walk(directory):
        for filenames in filenames:
            if filenames.split('.')[-1].lower() == 'xls':
                #print("Старый файл:",  filenames)
                file = os.path.join(dirpath,filenames)
                mydoc = minidom.parse(file)
                items = mydoc.getElementsByTagName('Data')
                it = items[4].firstChild.data.split(':')[1].strip()
                #print(it)
                ren(filenames,it, dirpath)
                print(file)
if __name__ == '__main__':
    print_hi('PyCharm')
messagebox.showinfo("Title", "Добавил период")