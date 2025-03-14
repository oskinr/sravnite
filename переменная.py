from tkinter import *
from tkinter  import messagebox
from tkinter.ttk import Combobox


def show_values1():
    global a

    a = combo.get()
    if a == '':
      a = "0"
    print(a)





root= Tk()
root.geometry("220x60")



bu1=Button(text='Enter', command = show_values1)
bu1.place(x = 5, y = 1)

#комбобоксы для ввода ключа слияния и сравнения
combo = Combobox (root, values='номер')
combo.place(x = 50, y = 1)



root.mainloop()
