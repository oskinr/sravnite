from tkinter import *
from tkinter  import messagebox
from tkinter.ttk import Combobox
import tkinter as tk

#создаем счетчик кликов
def foo():
  counter = 0
  def bar():
    nonlocal counter

    counter += 1
    label1.configure(text=counter)
  return bar
bar = foo()




#это функция для вставки нуля если о окне ввода 0 символов
def show_values():
    global rows
    rows = combo.get()
    if len(rows) == 0:
      rows = "0"
    print(rows)
    label3.configure(text=rows)

#Удаление текста из Меток label ов
def remove_text():
    label3.config(text="")
    label1.config(text="")



#это само окно
root= Tk()
root.geometry("230x90")



bu1=Button(text='Enter', command =show_values )
bu1.place(x = 5, y = 1)
bu2=Button(text='delet', command =remove_text )
bu2.place(x = 5, y = 30)
bu3=Button(text='count', command =bar)
bu3.place(x = 5, y = 60)
#комбобоксы для ввода ключа слияния и сравнения
combo = Combobox (root, values='номер')
combo.place(x = 50, y = 1)

#текстовой вывод 1 фала
label3 = Label(root, text="", justify=tk.LEFT) # создаем текстовую метку
label3.place(x = 50, y = 25)

#текстовой вывод пути к  фалам
label1 = Label(root, text="",font="system") # создаем текстовую метку
label1.place(x = 55, y = 45)

#повтор запуска окна
root.mainloop()
