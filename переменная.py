from tkinter import *
from tkinter  import messagebox
from tkinter.ttk import Combobox


def show_values1():
    global rows

    rows = combo.get()
    if rows == '':
      rows = "0"
    print(rows)





root= Tk()
root.geometry("220x60")



bu1=Button(text='Enter', command = show_values1)
bu1.place(x = 5, y = 1)

#комбобоксы для ввода
combo = Combobox (root, values=[0,6])
#Можно без условия и функций сразу из комбобокса возвращать get ом значение 0 или 6
# combo.current(0)#номер значения по порядку
# rows = combo.get() #Возвращает значение в rows по номеру ноль по порядку current из комбобокса
combo.place(x = 50, y = 1)



root.mainloop()
