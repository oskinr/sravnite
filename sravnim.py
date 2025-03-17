from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import pandas as pd
import tkinter as tk
from tkinter.ttk import Combobox




# Открываем файл номер один
def openanyfile():
    try:
       global selected_file
       global filename1
       selected_file  = filedialog.askopenfilename()
       label1.configure(text= selected_file.split('/')[-1])
       filename1 = selected_file
       df1 = pd.read_excel(filename1)
       print(df1)
    except Exception as e:
       messagebox.showerror(title="ошибка", message="Привет от системы на Файл 1 : " + str(e))




# Открываем файл 2
def openanyfile2():
    try:
        global selected_file2
        selected_file2  = filedialog.askopenfilename()
        label2.configure(text=selected_file2.split('/')[-1])
        filename2 = selected_file2
        df2 = pd.read_excel(filename2)
        #df2 = pd.read_excel(filename2,nrows = 7)
        print(df2)
    except Exception as e:
       messagebox.showerror(title="ошибка", message="Система ругается на Файл 2 : " + str(e))




#Читаем файл 1
def showfile1():
    #global df1
    filename1 = selected_file
    df1 = pd.read_excel(filename1)
    #df1.columns.tolist()  #df2.columns.tolist(['слияние'])
    #df1 = pd.read_excel(filename1, usecols='A, B ,C ,D, F',nrows = 3)
    label3.configure(text=df1.columns.tolist())
    #label3.configure(text= df1.keys().tolist())
    col_name = list(df1.columns)
    combo['values'] = col_name

    # new_option = df1.columns.tolist()[0]
    # combo['values'] = tuple(list(combo['values']) + [new_option])
    # combo.current(len(combo['values'])+1)

#Читаем файл 2 правка Федор
def showfile2():
    #global df2
    filename2 = selected_file2
    df2 = pd.read_excel(filename2)
    # df2 = pd.read_excel(filename2,usecols='D ,E, L',nrows = 3)
    label4.configure(text=df2.columns.tolist())
    dfr = pd.DataFrame(df2)
    col_name = list(dfr.columns)
    combo2['values'] = col_name




#возвращаем из окон entry значения в переменную
def show_message():
    global st1,st2
    st1 = combo.get()
    st2 = combo2.get()
    # st1 = entry.get()
    # st2 = entry2.get()
    # окно для заголовка Слияния
    if st1 != '':
        messagebox.showinfo(title = 'Заголовок слияния', message = st1 )
    else:
        messagebox.showerror(title="ошибка", message='Не введен заголовок слияния')

    # окно для заголовка сравнения
    if st2 != '':
        messagebox.showinfo(title = 'Заголовок сравнения', message = st2)
    else:
        messagebox.showerror(title="ошибка", message='Не введен заголовок для сравнения')

    try:
# файл из скрипта выгрузка из АСУ
        filename1 = selected_file
        df1 = pd.read_excel(filename1)
        df1 = df1.rename(columns={st2: 'Сравниваем', st1: 'Слияние'})
                             #usecols=['Улица','HOUSE_NOMER','Квартира','VAL_STR', 'ALL_SQR'])
                             #usecols='A, B ,C ,D, F')# файл ЕГРН
    except Exception as e:
       messagebox.showerror(title="ошибка", message="Система ругается на Файл 2 : " + str(e))

# файл из скрипта выгрузка из ЕГРН
    filename2=selected_file2
    df2 = pd.read_excel(filename2)
    df2 = df2.rename(columns={st1: 'Слияние',st2: 'Сравниваем'})
    print(df2)# usecols=['VAL_STR','ALL_SQR','Comment'])

    #Читаем ключи в датафрейм 1 проверяем
    #global a
    a = df1.keys().tolist()
    # print(a)
    # print(st1)
     # окно для проверки ключа слияния

    if 'Слияние' in a:
        messagebox.showinfo(title = 'Слияние', message ='Ключь для слияния создан' )
    else:

        messagebox.showwarning(title="ошибка", message='Вы ввели не верные заголовки, программа не может создать ключь для слияния, в обоих файлах должны быть одинаковые названия заголовков проверте и введите правильно')

    try:
# Сравниваем строки осуществляем слияние правое т.е к egrn прикрепим строчки из ас
        df3 = pd.merge(df1, df2, left_on=['Слияние'], right_on=['Слияние'], suffixes=('_Файл_1', '_Файл_2'),  how='right')
    except Exception as e:
        messagebox.showerror(title="ошибка", message="Система : " + str(e))




#Сохраним в файл
    df3.to_excel('out.xlsx')
    messagebox.showinfo("Title", "Создан фал out.xlsx")
#Выведем таблицу с результатом сравнения на экран
    print(df3)
    label5.configure(text=df3)


    if label5.configure(text=df3) != '':
         messagebox.showinfo(title = 'слияние', message='Поздравляю! Все прошло успешно' )
    else:
        messagebox.showwarning(title="ошибка", message='Не верно переименновали в файлах')
#Удаление текста из Меток label ов
def remove_text():
    label3.config(text="")
    label4.config(text="")
    label5.config(text="")








#Выведем таблицу с результатом сравнения на экран
window = Tk()
window.title("Сравнить файлы")
window.geometry("1500x500")
#window.iconbitmap(default="boss.ico")
window.iconphoto(False, tk.PhotoImage(file='tools.png'))
#window.iconbitmap(default="boss.ico")
frame = Frame(window)
#frame.pack(expand=True)
frame.pack(fill=Y)
#frame = Frame(master=window, relief=GROOVE, borderwidth=5)




# создаем текстовую метку
label = Label(frame, text="Выбери файл 1 побольше, затем файл 2 меньше," "\n"
                          "предварительно сделав там одинаковые названия" '\n'
                          "столбиков например номер и площадь в обоих" '\n'
                          "файлах, слияние будет по первому столбику номер" '\n'
                          "потом сможете сравнить площадь")
label.grid(row=0, column=1, pady=5)

# method_lbl = Label(frame, text="Выберите файлы")
# method_lbl.grid(row=0, column=1)


#подпись для поля ввода 1:
base_lbl = Label(frame,text="Введите заголовок для слияния")
base_lbl.grid(row=2, column=1, pady=1)
#поля для ввода значений 1:
# entry = Entry(frame)
# entry.grid(row=2, column=2, pady=2)

#подпись для поля ввода 2:
height_lbl = Label(frame,text="Введите заголовок для сравнения")
height_lbl.grid(row=3, column=1, pady=1)
#поле для ввода 2:
# entry2 = Entry(frame)
# entry2.grid(row=3, column=2, pady=2)

#установить фокус на ввод текста
#entry.focus()

#кнопки
file1 = Button(frame, text="Файл 1",command=openanyfile)
file1.grid(row=2, column=0)

file2 = Button(frame, text="Файл 2",command=openanyfile2)
file2.grid(row=3, column=0)

calc_btn = Button(frame, text="Сохранить в отдельный файл",command=show_message)
calc_btn.grid(row=6, column=1)

show = Button(frame, text="Заголовки файла 1",command=showfile1)
show.grid(row=1, column=1)

show2 = Button(frame, text="Заголовки файла 2",command=showfile2)
show2.grid(row=1, column=2)

show3 = Button(frame, text="Удалить текст",command=remove_text)
show3.grid(row=1, column=3)

#текстовой вывод пути к  фалам
label1 = Label(frame, text="",font="system") # создаем текстовую метку
label1.grid(row=2, column=6, pady=10)

label2 = Label(frame, text="",font="system") # создаем текстовую метку
label2.grid(row=3, column=6, pady=10)

#вывод 3файлов в  текст
#текстовой вывод 1 фала
label3 = Label(frame, text="", justify=tk.LEFT) # создаем текстовую метку
label3.grid(row=2, column=7, pady=10)
#текстовой вывод 2 фала
label4 = Label(frame, text="", justify=tk.LEFT) # создаем текстовую метку
label4.grid(row=3, column=7, pady=10)
#текстовой вывод сравнения фалов
label5 = Label(frame, text="", justify=tk.LEFT)#, bg='aquamarine') # создаем текстовую метку
label5.grid(row=1, column=7, pady=10)


#комбобоксы для ввода ключа слияния и сравнения
combo = Combobox (frame, values='номер')
combo.grid(row=2, column=2, pady=10)

combo2 = Combobox (frame, values='площадь')
combo2.grid(row=3, column=2, pady=10)

window.mainloop()
