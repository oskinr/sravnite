from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import pandas as pd
import tkinter as tk
from tkinter.ttk import Combobox



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
    # global df1
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



    try:
        # файл из скрипта выгрузка из АСУ
        global rows

        df1 = pd.read_excel(selected_file,skiprows=int(rows))
        df1 = df1.rename(columns={st2: 'Сравниваем', st1: 'Слияние'})
        # usecols=['Улица','HOUSE_NOMER','Квартира','VAL_STR', 'ALL_SQR'])
        # usecols='A, B ,C ,D, F')# файл ЕГРН
    except Exception as err:
        messagebox.showerror(
            title="ошибка", message="🔒 Привет от системы : " + str(err))

        print

# файл из скрипта выгрузка из ЕГРН
    try:

        df2 = pd.read_excel(selected_file2, skiprows=int(rows))
        df2 = df2.rename(columns={st1: 'Слияние', st2: 'Сравниваем'})
        print(df2)  # usecols=['VAL_STR','ALL_SQR','Comment'])
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
        # Сравниваем строки осуществляем слияние правое т.е к egrn прикрепим строчки из асу
        df3 = pd.merge(df1, df2, left_on=['Слияние'], right_on=[
                       'Слияние'], suffixes=('_Файл_1', '_Файл_2'),  how='right')
    except Exception as err:
        messagebox.showerror(
            title="ошибка", message="🔒 Cистема не верный столбик для сравнения - его нет в файле : " + str(err))

    try:
        # Сохраним в файл
        b = df3.to_excel('out.xlsx')
        messagebox.showinfo("Title", "Создан фал out.xlsx")
    except Exception as err:
        messagebox.showerror(
            title="ошибка", message="🔒 Cистема записать в файл out.xlsx не удалось возможно он открыт - закройте : " + str(err))
# Выведем таблицу с результатом сравнения на экран
    print(df3)
    label5.configure(text=df3)

    if b != '':
        messagebox.showinfo(
            title='слияние', message='Поздравляю! Все прошло успешно')
    else:
        messagebox.showwarning(
            title="ошибка", message='Не совпадают заголовки в файлах')
# Удаление текста из Меток label ов


def remove_text():
    label3.config(text="")
    label4.config(text="")
    label5.config(text="")


# Выведем таблицу с результатом сравнения на экран
window = Tk()
window.title("Сравнить файлы")
window.geometry("1500x500")


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

# method_lbl = Label(frame, text="Выберите файлы")
# method_lbl.grid(row=0, column=1)


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
file1 = Button(frame, text="Файл 1", command=openanyfile)
file1.grid(row=2, column=0)

file2 = Button(frame, text="Файл 2", command=openanyfile2)
file2.grid(row=3, column=0)

calc_btn = Button(frame, text="Сохранить в отдельный файл", command=show_message)
calc_btn.grid(row=6, column=1)

calc_btn = Button(frame, text="удалить строки f1", command=showrows)
calc_btn.grid(row=6, column=2)

calc_btn = Button(frame, text="удалить строки f2", command=showrows2)
calc_btn.grid(row=6, column=3)

show = Button(frame, text="Заголовки файла 1", command=showfile1)
show.grid(row=1, column=1)

show2 = Button(frame, text="Заголовки файла 2", command=showfile2)
show2.grid(row=1, column=2)

show3 = Button(frame, text="Удалить текст", command=remove_text)
show3.grid(row=1, column=3)

# текстовой вывод пути к  фалам
label1 = Label(frame, text="", font="system")  # создаем текстовую метку
label1.grid(row=2, column=6, pady=10)

label2 = Label(frame, text="", font="system")  # создаем текстовую метку
label2.grid(row=3, column=6, pady=10)

# вывод 3файлов в  текст
# текстовой вывод 1 фала
label3 = Label(frame, text="", justify=tk.LEFT)  # создаем текстовую метку
label3.grid(row=2, column=7, pady=10)
# текстовой вывод 2 фала
label4 = Label(frame, text="", justify=tk.LEFT)  # создаем текстовую метку
label4.grid(row=3, column=7, pady=10)
# текстовой вывод сравнения фалов
# , bg='aquamarine') # создаем текстовую метку
label5 = Label(frame, text="", justify=tk.LEFT)
label5.grid(row=1, column=7, pady=10)


# комбобоксы для ввода ключа слияния и сравнения
combo = Combobox(frame, values='номер')
combo.grid(row=2, column=2, pady=10)

combo2 = Combobox(frame, values='площадь')
combo2.grid(row=3, column=2, pady=10)

combo3 = Combobox(frame, values=[0,6])
combo3.current(0)
rows = combo3.get()
combo3.grid(row=4, column=2, pady=10)
# text = Text(frame, width=25, height=5, bg='white', fg='grey', wrap=WORD)
# text.grid(row=0, column=2, pady=5)

window.mainloop()
