import os
from tkinter import filedialog
from xml.dom import minidom
from tkinter import * 
from tkinter import ttk
from tkinter import filedialog as fd


        

def selectDir():
    global file,items,nom,telo,directory,dirpath,it
    directory = filedialog.askdirectory()+'/'
    label1.configure(text = directory)
    for dirpath, dirnames, filenames in os.walk(directory):
        for filenames in filenames:
            if filenames.split('.')[-1].lower() == 'xls':
         
                file = os.path.join(dirpath,filenames)
                mydoc = minidom.parse(file)
                items = mydoc.getElementsByTagName('Data')
                rows = combobox.get()
                it = items[int(rows)].firstChild.data.split(':')[1].strip()
                column_listbox.insert(0,it)
   
def per():   
    for dirpath, dirnames, filenames in os.walk(directory):
        ig = 0
        for filenames in filenames:
                
            if filenames.split('.')[-1].lower() == 'xls':
                
            # newfile = file.split('_')[0].strip() + it + '_' + '.xls'
            # os.rename(os.path.join(dirpath, file), os.path.join(dirpath, newfile))   
            #print(file)
                file = os.path.join(dirpath,filenames)
                mydoc = minidom.parse(file)
                items = mydoc.getElementsByTagName('Data')
                rows = combobox.get()
            
                it = items[int(rows)].firstChild.data.split(':')[1].strip()
                
                column_listbox.insert(0,it)
                ig += 1
                #per(filenames, it,ig, dirpath)
            #print(it)
            # nom, dirpath= file.split('_')
            # print(rows)
                newfile = it + '_' + filenames.split('_')[0].strip() +'_' + str(ig) + '.xls'
                os.rename(os.path.join(dirpath, filenames), os.path.join(dirpath, newfile))
        
                #print('Новый файл:', newfile)
                #print(ig)

          

    #files=sorted([path for path in os.listdir(directory) if os.path.isfile(directory+path)])


   
    # i = 1
    # dir = directory
    # for file in os.listdir(dir):
    #     #if file.split('.')[-1].lower() == 'xls':
    #     if file.endswith('xls'):
    #         print(it)
    #         os.rename(f'{dir}/{file}', f'{dir}/{0,it}_отчет_{i}_19.07.{'xls'}')
    #         i = i + 1


    # newfile = it + '_' + file.split('_')[0].strip() + '.xls'
    # os.rename(os.path.join(dirpath, file), os.path.join(dirpath, newfile))
    # #print('Новый файл:', newfile)
    # print(file)
        
           

    
    
        # if file.split('.')[-1].lower() == 'xls':
        #     os.rename(file, it + '_' + nom + '_' + telo +'.xls')
  
def del_list():
    select = list(column_listbox.curselection())
    select.reverse()
    for i in select:
        column_listbox.delete(i)          





def spisok():
    global it
    rows = combobox.get()
    it = items[int(rows)].firstChild.data.split(':')[1].strip()
    column_listbox.insert(0,it)
    print(rows)
    
def add_item():
    column_listbox.insert(END, combo4.get())
    combo4.delete(0, END)
    combo4.insert(0,it)
    #directory  = filedialog.askdirectory()
    # os.chdir(directory)
    # text = os.listdir()
    # for folder in os.listdir():
    #     if os.path.isdir(folder):
    #      день, месяц, год = folder.split('.')
    #      os.rename(folder, день + '.' + месяц + '.' + год)
    
    # #words = text.split('/', 1)
    # print( folder, день, месяц, год)
    
    
    
    # for dirpath, dirnames, filenames in os.walk(directory):
    #     for filenames in filenames:
    #         if filenames.split('.')[-1].lower() == 'xls':
                
                
                
    #             file = os.path.join(directory,filenames)
    #             print(filenames)
                
                
    #             mydoc = minidom.parse(file)
    #             items = mydoc.getElementsByTagName('Data')
    #             it = items[4].firstChild.data.split(':')[1].strip()
    #             selectDir(filenames,it, dirpath)
                
    #             #print(it)
                
    #             column_listbox.insert(0,it)
    #             label1.configure(text= it)
                

window = Tk()
window.title("Сравнить файлы")
window.geometry("1000x400")


column_listbox = Listbox(selectmode=EXTENDED)
column_listbox.grid(row=1, column=1, columnspan=2, sticky=NSEW, padx=5, pady=5)

label1 = Label(window, text="",font="system") # создаем текстовую метку
label1.grid(column=3, row=1, padx=6, pady=6)              

Button(text="Найти период", command=spisok).grid(row=0, column=10, columnspan=2, sticky=NSEW, padx=5, pady=5)
Button(text="Папка", command=selectDir).grid(row=2, column=10, columnspan=2, sticky=NSEW, padx=5, pady=5)
Button(text="Удалить", command=del_list).grid(row=3, column=10, columnspan=2, sticky=NSEW, padx=5, pady=5)
Button(text="Добавить", command=add_item).grid(row=4, column=10, columnspan=2, sticky=NSEW, padx=5, pady=5)
Button(text="Переименовать", command=per).grid(row=5, column=10, columnspan=2, sticky=NSEW, padx=5, pady=5)
combo4 = ttk.Combobox(values='')
combo4.grid(row=2, column=3, padx=6, pady=6)

combobox = ttk.Combobox( values=[0,6])
combobox.set(4)
#combobox.current(4)
rows = combobox.get()
combobox.grid(row=2, column=1, padx=6, pady=6)
window.mainloop()   
