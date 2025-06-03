import os
import win32com.client
from tkinter import filedialog
from tkinter import *
from tkinter.messagebox import showerror, showinfo, askyesno
import tkinter as tk
import pathlib
from tkinter import ttk


appdir = pathlib.Path(__file__).parent.resolve()

def del_columns():
    def selectDir():
        global directory  # Используем локальную переменную вместо глобальной
        try:
            directory = filedialog.askdirectory()
            files = os.listdir(directory)
            files = [fname for fname in files if fname.endswith(('.xls', '.xlsx', 'xlsm'))]
            label1.configure(text=directory.split('/')[-1])
            result = askyesno(title="Проверяем файлы в папке",
                              message=f"В этих файлах будем удалять столбцы в папке:\n\n{'\n'.join(files)}")
            if result:
                showinfo("Результат", "Файлы выбраны, выберите столбец и нажмите удалить")
            else:
                showinfo("Результат", "Операция отменена")
        except Exception as err:
            showerror(title="Ошибка", message=f"Система: {err}")

    def del_list():
        try:
            path = directory
            exts = ["xls", "xlsx", "xlsm"]  # Список поддерживаемых расширений
            processed_files = []  # Здесь будут храниться имена обработанных файлов
            selected_column = column_combobox.get().strip()  # Выбранный столбец
            if not selected_column:
                raise ValueError("Необходимо выбрать столбец для удаления.")

            excel_app = win32com.client.Dispatch("Excel.Application")  # Создаем экземпляр Excel
            excel_app.Visible = False  # Скрываем окно приложения
            excel_app.DisplayAlerts = False  # Отключаем уведомления

            # Если каталог существует, обрабатываем каждый файл
            if os.path.exists(path):
                files_in_dir = os.listdir(path)
                for filename in files_in_dir:
                    full_path = os.path.join(path, filename)
                    ext = os.path.splitext(filename)[1][1:]
                    if ext.lower() in exts:
                        workbook = excel_app.Workbooks.Open(full_path)
                        sheet = workbook.ActiveSheet
                        sheet.Columns(selected_column).Delete()  # Удаляем указанный столбец
                        workbook.Save()  # Сохраняем изменения
                        workbook.Close()  # Закрываем книгу
                        processed_files.append(filename)  # Добавляем имя файла в список обработанных
            else:
                showerror("Результат", f"Путь {path} не существует.")
        except Exception as err:
            showerror(title="Ошибка", message=f"Система: Названия столбиков должны быть большими буквами и в английской раскладке.\n{err}")
        finally:
            excel_app.Quit()  # Завершаем приложение Excel
            del excel_app  # Освобождаем ресурсы

        if len(processed_files) > 0:
            showinfo("Результат", "Файлы обработаны")
            for fname in processed_files:
                column_listbox.insert(0, fname)
        else:
            showerror("Результат", "Ни одного подходящего файла не найдено.")

    # Основной интерфейс окна редактирования XML
    form = Toplevel()
    form.title("Редактирование Excel-файлов")
    form.geometry("530x250+950+100")
    form.iconbitmap(os.path.join(appdir, 'osa.ico'))
    form.attributes('-topmost', True)

    column_listbox = tk.Listbox(form, selectmode=tk.SINGLE)
    column_listbox.grid(row=1, column=1, columnspan=2, sticky=tk.NSEW, padx=5, pady=5)

    label1 = tk.Label(form, text="", font="system")
    label1.grid(row=0, column=10, columnspan=2, sticky=tk.NSEW, padx=5, pady=5)

    Label = tk.Label(form, text="Выберите папку и столбец для удаления:")
    Label.grid(row=0, column=2, padx=5, pady=5)

    columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
    column_combobox = ttk.Combobox(form, values=columns, state="normal")
    column_combobox.current(1)
    column_combobox.grid(row=0, column=1, padx=5, pady=5)

    Button(form, text="Открыть папку", command=selectDir).grid(row=1, column=10, columnspan=2, sticky=tk.NSEW, padx=5, pady=5)
    Button(form, text="Удалить", command=del_list).grid(row=3, column=10, columnspan=2, sticky=tk.NSEW, padx=5, pady=5)

    form.mainloop()