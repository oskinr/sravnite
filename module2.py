import os
from tkinter import filedialog
from xml.dom import minidom
from tkinter import * 
from tkinter import ttk
import tkinter as tk
import pathlib,os.path
appdir = pathlib.Path(__file__).parent.resolve()
def some_form():    

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
       
                    file = os.path.join(dirpath,filenames)
                    mydoc = minidom.parse(file)
                    items = mydoc.getElementsByTagName('Data')
                    rows = combobox.get()
            
                    it = items[int(rows)].firstChild.data.split(':')[1].strip()
                
                    column_listbox.insert(0,it)
                    ig += 1
          
                    newfile = it + '_' + filenames.split('_')[0].strip() +'_' + str(ig) + '.xls'
                    os.rename(os.path.join(dirpath, filenames), os.path.join(dirpath, newfile))

  
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

                
  # Чтобы дочернее окно было поверх родительского
   
    form =Tk()
    
    form.title("Переименовать файлы")
    # Устанавливаем размеры и позицию окна
# Например, размещаем окно размером 400х300 пикселей в позиции (200, 150)
    form.geometry("550x400+700+150")
    form.iconbitmap(os.path.join(appdir,'osa.ico'))
    form.lift()
    form.attributes('-topmost', True)
    
    # Настройки для улучшения размера и центровки
    form.columnconfigure(0, weight=0)
    form.rowconfigure(0, weight=0)

# Лист бокс (список)
    column_listbox = tk.Listbox(form, selectmode=tk.EXTENDED)
    column_listbox.grid(row=1, column=1, columnspan=2, sticky=tk.NSEW, padx=5, pady=5)

# # Пример заполнения списка элементами
#     for i in range(10):
#         column_listbox.insert(i, f"Элемент {i}")

    # column_listbox = Listbox(form, selectmode=EXTENDED)
    # column_listbox.grid(row=1, column=1, columnspan=2, sticky=NSEW, padx=5, pady=5)

    label1 = Label(form, text="",font="system") # создаем текстовую метку
    label1.grid(column=3, row=1, padx=6, pady=6)              

    Button(form,text="Найти период", command=spisok).grid(row=0, column=10, columnspan=2, sticky=NSEW, padx=5, pady=5)
    Button(form,text="Папка", command=selectDir).grid(row=2, column=10, columnspan=2, sticky=NSEW, padx=5, pady=5)
    Button(form,text="Удалить", command=del_list).grid(row=3, column=10, columnspan=2, sticky=NSEW, padx=5, pady=5)
    Button(form,text="Добавить", command=add_item).grid(row=4, column=10, columnspan=2, sticky=NSEW, padx=5, pady=5)
    Button(form,text="Переименовать", command=per).grid(row=5, column=10, columnspan=2, sticky=NSEW, padx=5, pady=5)
    combo4 = ttk.Combobox(form,values='')
    combo4.grid(row=2, column=3, padx=6, pady=6)

    combobox = ttk.Combobox(form, values=[0,6])
    combobox.set(4)
#combobox.current(4)
    rows = combobox.get()
    combobox.grid(row=2, column=1, padx=6, pady=6)
    form.mainloop() 
