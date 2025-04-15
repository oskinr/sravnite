from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter.ttk import *
from time import sleep
import os
import zipfile
from pathlib import PurePath
import sys
import codecs
import win32com.client
from tkinter.messagebox import showinfo, askyesno

sys.stdout = codecs.getwriter("utf-8")(sys.stdout.detach())
# Открываем файл 1
def openanyfile():
    try:
        global selected_file
        selected_file = filedialog.askopenfilename()
        label1.configure(text=selected_file.split('/')[-1])
        df1 = pd.read_excel(selected_file)
    except Exception as err:
        messagebox.showerror(
            title="ошибка", message="🔒 Привет от системы, что то с Файл 1 формат xlsx? : " + str(err))

# Открываем файл 2
def openanyfile2():
    try:
        global selected_file2
        selected_file2 = filedialog.askopenfilename()
        label2.configure(text=selected_file2.split('/')[-1])
        df2 = pd.read_excel(selected_file2)
    except Exception as err:
        messagebox.showerror(
            title="ошибка", message="🔒 Привет от системы, что то с Файл 2 : " + str(err))


# Читаем файл 1
def showfile1():
    #global df1
    df1 = pd.read_excel(selected_file)
    label3.configure(text=df1.keys().tolist())
    col_name = list(df1.columns)
    combo['values'] = col_name

# Читаем файл 2
def showfile2():
    # global df2
    df2 = pd.read_excel(selected_file2)
    label4.configure(text=df2.columns.tolist())
    col_name = list(df2.columns)
    combo2['values'] = col_name


# добавить сколько строк пропустить 1
def showrows():
    global rows
    rows = combo3.get()
    df1 = pd.read_excel(selected_file, skiprows=int(rows))
    label3.configure(text=df1.keys().tolist())
    col_name = list(df1.columns)
    combo['values'] = col_name

# добавить сколько строк пропустить 2
def showrows2():
    #global rows
    rows = combo3.get()
    df2 = pd.read_excel(selected_file2, skiprows=int(rows))
    label4.configure(text=df2.keys().tolist())
    col_name = list(df2.columns)
    combo2['values'] = col_name



# возвращаем из окон entry значения в переменную
def show_message():
    global st1, st2
    st1 = combo.get()
    st2 = combo2.get()

# сообщение для заголовка Слияния
    if st1 != '':
        messagebox.showinfo(title='Заголовок слияния', message=st1)
    else:
        messagebox.showerror(
            title="ошибка", message='Не введен заголовок слияния')

    # сообщение для заголовка сравнения
    if st2 != '':
        messagebox.showinfo(title='Заголовок сравнения', message=st2)
    else:
        messagebox.showerror(
            title="ошибка", message='Не введен заголовок для сравнения')


# файл из скрипта выгрузка из АСУ
    try:
        df1 = pd.read_excel(selected_file,skiprows=int(rows))
        df1 = df1.rename(columns={st2: 'Сравниваем', st1: 'Слияние'})
    except Exception as err:
        messagebox.showerror(
            title="ошибка", message="🔒 Привет от системы : " + str(err))


# файл из скрипта выгрузка из ЕГРН
    try:
        df2 = pd.read_excel(selected_file2, skiprows=int(rows))
        df2 = df2.rename(columns={st1: 'Слияние', st2: 'Сравниваем'})
    except Exception as err:
        messagebox.showerror(
            title="ошибка", message="🔒 Привет от системы : " + str(err))
    # Читаем ключи в датафрейм 1 проверяем
    # global a
    key_slianie = df1.keys().tolist()

    # print(a)
    # print(st1)




# Сообщение для проверки ключа слияния
    if 'Слияние' in key_slianie:
        messagebox.showinfo(title='Слияние', message='Ключ для слияния создан')

    else:
        messagebox.showwarning(
            title="ошибка", message='Вы ввели не верные заголовки, программа не может создать ключь для слияния, в обоих файлах должны быть одинаковые названия заголовков проверте и введите правильно')


    try:
        global df3
    # Сравниваем строки осуществляем слияние правое т.е к egrn прикрепим строчки из асу
        df3 = pd.merge(df1, df2, left_on=['Слияние'], right_on=['Слияние'], suffixes=('_Файл_1', '_Файл_2'),  how='right')
    except Exception as err:
        messagebox.showerror(
            title="ошибка", message="🔒 Cистема не верный столбик для сравнения - его нет в файле : " + str(err))

    try:
        bar = df3.style.set_properties(**{'border': '1.3px solid black', 'color': 'black'}).to_excel('out.xlsx', index=False)
        # Сохраним в файл
        #b = df3.to_excel('out.xlsx')
        #запуск прогрессбара и счетчик % для него если дата фрейм не пустой
        if bar !='':
         for i in range(number):
            progressbar.configure(value= i / (number / 101))
            label6.configure(text = f'{int(i / (number / 101))} %' )
            sleep(0.01)
            progressbar.update()


        messagebox.showinfo("Title", "Создан фал out.xlsx")
    except Exception as err:
        messagebox.showerror(
            title="ошибка", message="🔒 Cистема записать в файл out.xlsx не удалось возможно он открыт - закройте : " + str(err))

# Выведем таблицу с результатом сравнения на экран

    label5.configure(text=df3)
    col_name = list(df3.columns)
    combo4['values'] = col_name

    if bar != '':
        messagebox.showinfo(
        title='слияние', message='Поздравляю! Все прошло успешно')
        #progressbar.stop()      # останавливаем progressbar

    else:
        messagebox.showwarning(
            title="ошибка", message='Не совпадают заголовки в файлах')

# Удаление текста из Меток label ов
def remove_text():
    label1.config(text="")
    label2.config(text="")
    label3.config(text="")
    label4.config(text="")
    label5.config(text="")
    label7.config(text="")

def highlight_col(x):
    #copy df to new - original data are not changed
    df = x.copy()
    #set by condition
    mask = df['compare'] == False
    df.loc[mask, :] = 'background-color: yellow'
    df.loc[~mask,:] = 'background-color: white'
    return df

def add_item():
    lbox.insert(END, combo4.get())
    combo4.delete(0, END)

def del_list():
    select = list(lbox.curselection())
    select.reverse()
    for i in select:
        lbox.delete(i)

def del_tree():
    # selected_item = tree_view.selection()[0] # get selected item
    # tree_view.delete(selected_item)
    x = tree_view.get_children()
    for item in x:
        tree_view.delete(item)

def print_list():
    df = (lbox.get(0, END))
    df = " ".join(lbox.get(0, END))
    modified_list = (df.split())
    try:
      df4 =  df3[modified_list]
      # сравниваем столбики  и записываем результат сравнения в compare
      df = df4
      df4['compare'] = df4['Сравниваем_Файл_1'] == df4['Сравниваем_Файл_2']


      df4.style.apply(highlight_col, axis=None).set_properties(**{'border': '1.3px solid grey','color': 'black'}).to_excel('outfinish.xlsx', index=False)
      #df4.style.apply(highlight_col, axis=None).to_excel('outfinish.xlsx', index=False)
      #df4.style.set_properties(**{'border': '1.3px solid black', 'color': 'black'}).to_excel('outfinish.xlsx', index=False)
      #df4.to_excel('outfinish.xlsx', index=False)
      messagebox.showinfo("Title", "Создан фал outfinish.xlsx")
    except Exception as err:
        messagebox.showerror(
            title="ошибка", message="🔒 Система : " + str(err))



def click():
    global directory
    #Получаем список файлов в директории/каталоге os.listdir(directory)
    directory  = filedialog.askdirectory(**options)
    files = os.listdir(directory)
    result = askyesno(title="Подтвержение операции", message=("Файлы в папке:\n\n" + "\n".join(files)),)
    if result:
      zip_ex()
    else:
        showinfo("Результат", "Операция отменена")




def zip_ex():
    if directory:
        current_dir.set(directory)
        list_files(directory)

def list_files(directory):
        for filename1 in os.listdir(directory):
            if os.path.isfile(os.path.join(directory, filename1)):
                tree_view.insert("", "end", values=(filename1,))

            #for file in os.listdir(directory):
                filename = os.fsdecode(filename1)
                path = os.path.join(directory, filename)
                #print(path)

                if filename.endswith('.zip'):
                    try:
                        with zipfile.ZipFile(path) as zf:
                            filik = zf.namelist()
                            namefaile = filik[0]
                            old_file = f'{directory}\\{namefaile}'
                            new_file = f'{directory}\\{PurePath(filename).stem}{".xls"}'
                            zf.extract(namefaile, directory)
                    #messagebox.showinfo("извлек", path)
                    except zipfile.BadZipFile as error:
                        messagebox.showerror("ошибка", error)
                    if os.path.exists(new_file):
                        os.remove(new_file)
                        os.rename(old_file, new_file)
                        print(f"из {filename} извлечен файл:{os.path.basename(new_file)}")
                        #tree_view.insert(f"из {filename} извлечен файл:{os.path.basename(new_file)}")
                    else:
                        os.rename(old_file, new_file)

                    label7.configure(text=f" Из:{filename}\n Извлечен файл :\n {os.path.basename(new_file)}")
def convert():
    try:  
        file = filedialog.askopenfilename().replace('/', '\\')
        #print(file)
        wbf = file + "x"
        #print("конвертация xml")

        office = win32com.client.Dispatch("Excel.Application")
        wb = office.Workbooks.Open(file)
        sheet = wb.ActiveSheet
        office.DisplayAlerts = False  # не спрашивает перезаписать файл
    # val = sheet.Cells(1,1)
        num = [r for r in sheet.Range("A8:K8")]
        print(*num)
        
        wb.SaveAs(Filename=wbf, FileFormat=51)
        wb.Close(True)
        office.Quit()
        messagebox.showinfo("Конвертация xml в xlsx",wbf )
    except Exception as err:
        messagebox.showerror(
            title="ошибка", message="🔒 Система : " + str(err))
def per():    
    subprocess.run(['python', 'pereimenovat.py'])  # "запускаем другой скрипт" 


window = Tk()
number = 284
window.title("Сравнить файлы")
window.geometry("1500x700")



options = {"initialdir": "/Downloads","title": "Выбери папку с архивами для разархивирования",
           "mustexist": True,"parent": window}


# window.iconbitmap(default="boss.ico")
window.iconphoto(False, tk.PhotoImage(file='osa.png'))
# window.iconbitmap(default="boss.ico")

# Отступ от верха окна
frame = Frame(window, width=400, height=100)
frame = Frame(window)
frame.pack(expand=False)
# frame.pack(fill=Y)
# frame = Frame(master=window, relief=GROOVE, borderwidth=5)


# создаем текстовую метку
label = Label(frame, text="Выбери файл 1 побольше, затем файл 2 меньше," "\n"
                          "предварительно сделав там одинаковые названия" '\n'
                          "столбиков например номер и площадь в обоих" '\n'
                          "файлах, слияние будет по первому столбику номер" '\n'
                          "потом сможете сравнить площадь")
label.grid(row=0, column=1, pady=5)



# подпись для поля ввода 1:
base_lbl = Label(frame, text="Введите заголовок для слияния")
base_lbl.grid(row=2, column=1, pady=1)
# поля для ввода значений 1:
entry = Entry(frame)
entry.grid(row=2, column=2, pady=2)

# подпись для поля ввода 2:
height_lbl = Label(frame, text="Введите заголовок для сравнения")
height_lbl.grid(row=3, column=1, pady=1)


height_lbl = Label(frame, text="Введите сколько строк пропустить")
height_lbl.grid(row=4, column=1, pady=1)

# поле для ввода 2:
# entry2 = Entry(frame)
# entry2.grid(row=3, column=2, pady=2)

# установить фокус на ввод текста
# entry.focus()

# кнопки
file1 = ttk.Button(frame, text="Файл 1", command=openanyfile)
file1.grid(row=2, column=0)

file2 = ttk.Button(frame, text="Файл 2", command=openanyfile2)
file2.grid(row=3, column=0)

calc_btn = ttk.Button(frame, text="Слияние в один файл", command=show_message)
calc_btn.grid(row=6, column=1)

calc_btn = ttk.Button(frame, text="удалить строки f1", command=showrows)
calc_btn.grid(row=6, column=2)

calc_btn = ttk.Button(frame, text="удалить строки f2", command=showrows2)
calc_btn.grid(row=6, column=3)

show = ttk.Button(frame, text="Показать заголовки f1", command=showfile1)
show.grid(row=1, column=1)

show2 = ttk.Button(frame, text="Показать заголовки f2", command=showfile2)
show2.grid(row=1, column=2)

show3 = ttk.Button(frame, text="Удалить текст", command=remove_text)
show3.grid(row=1, column=3)

# текстовой вывод пути к  фалам
label1 = ttk.Label(frame, text="", font="system")  # создаем текстовую метку
label1.grid(row=2, column=6, pady=10)

label2 = ttk.Label(frame, text="", font="system")  # создаем текстовую метку
label2.grid(row=3, column=6, pady=10)

# вывод 3файлов в  текст
# текстовой вывод 1 фала
label3 = ttk.Label(frame, text="", justify=tk.LEFT)  # создаем текстовую метку
label3.grid(row=2, column=7, pady=10)
# текстовой вывод 2 фала
label4 = ttk.Label(frame, text="", justify=tk.LEFT)  # создаем текстовую метку
label4.grid(row=3, column=7, pady=10)
# текстовой вывод сравнения фалов
# , bg='aquamarine') # создаем текстовую метку
label5 = ttk.Label(frame, text="", justify=tk.LEFT)
label5.grid(row=1, column=7, pady=10)

label6 = ttk.Label(text="0%", justify=tk.LEFT)
label6.pack(fill=X, padx=700, pady=5)

label7 = ttk.Label(text="", justify=tk.LEFT)
label7.pack(fill=X, padx=700, pady=5)

# комбобоксы для ввода ключа слияния и сравнения
combo = ttk.Combobox(frame, values='')
combo.grid(row=2, column=2, pady=10)

combo2 = ttk.Combobox(frame, values='')
combo2.grid(row=3, column=2, pady=10)

combo3 = ttk.Combobox(frame, values=[0,6])
combo3.current(0)
rows = combo3.get()
combo3.grid(row=4, column=2, pady=10)


progressbar = Progressbar(orient=HORIZONTAL, mode="determinate", length=500)
progressbar.pack(fill=X, padx=30, pady=5)

#блок листбокса
# label = ttk.Label(text='Собрать файл' )
# label.pack(fill=X, padx=310, pady=1)



lbox = Listbox(selectmode=EXTENDED)
lbox.pack(side=LEFT, padx=20, pady=20)

scroll = Scrollbar(command=lbox.yview)
scroll.pack(side=LEFT, fill=Y)

lbox.config(width=70, height=20, yscrollcommand=scroll.set)

f = Frame()
f.pack(side=LEFT, padx=10)


combo4 = Combobox(f, values='')
combo4.pack(fill=X, padx=90, pady=6)

Button(text="Разархивировать файлы - выбрать директорию", command=click).pack(fill=X, padx=90, pady=1)
Button(f, text="Добавить", command=add_item).pack(fill=X)
Button(f, text="Удалить", command=del_list).pack(fill=X)
Button(f, text="Собрать", command=print_list).pack(fill=X)
Button(f, text="Удалить список >>>", command=del_tree).pack(fill=X)
Button(f, text="Конвертировать", command=convert).pack(fill=X)
Button(f, text="Переименовать файлы в директории", command=per).pack(fill=X)

current_dir = tk.StringVar()

folder_label = tk.Label( textvariable=current_dir, font=("italic 14"))
folder_label.pack()

tree_view = ttk.Treeview( columns=("Files",), show="headings", selectmode="browse")
tree_view.heading("Files", text="Файлы в директории")
tree_view.pack(padx=20, pady=20, fill="both", expand=True)

#ttk.Button(text="Click", command=click).pack(anchor="center", expand=1)
window.mainloop()
