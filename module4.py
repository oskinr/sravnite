import os
from tkinter import *
from tkinter import filedialog, messagebox
import pyzipper
import tkinter as tk
import pathlib,os.path
appdir = pathlib.Path(__file__).parent.resolve()



def zip_arh():

    # Функция архивирования каждого файла отдельно
    def archive_each_file_separately(dir_path, password=None):
        xls_files = [f for f in os.listdir(dir_path) if f.endswith(('.xlsx', '.xls'))]
        
        if not xls_files:
            raise FileNotFoundError("Ни одного файла XLS не обнаружено.")
        
        for file in xls_files:
            full_path = os.path.join(dir_path, file)
            base_filename = os.path.splitext(file)[0]
            output_zip_path = os.path.join(dir_path, f'{base_filename}.zip')
            
            # Создаем архив для каждого файла
            if password is not None:
                with pyzipper.AESZipFile(output_zip_path, 'w', compression=pyzipper.ZIP_DEFLATED, encryption=pyzipper.WZ_AES) as zf:
                    zf.pwd = password.encode('utf-8')
                    zf.write(full_path, arcname=file)
            else:
                with pyzipper.AESZipFile(output_zip_path, 'w', compression=pyzipper.ZIP_DEFLATED, encryption=None) as zf:
                    zf.write(full_path, arcname=file)

    # Функция архивирования всех файлов в один архив
    def archive_all_files_together(dir_path, password=None):
        xls_files = [f for f in os.listdir(dir_path) if f.endswith(('.xlsx', '.xls'))]
        if not xls_files:
            raise FileNotFoundError("Ни одного файла XLS не обнаружено.")
        
        dir_name = os.path.basename(os.path.normpath(dir_path))
        output_zip_path = os.path.join(dir_path, f'{dir_name}_files.zip')
        
        # Список абсолютных путей файлов
        input_files = [os.path.join(dir_path, f) for f in xls_files]
        if password is not None:
        # Создание единого архива
            with pyzipper.AESZipFile(output_zip_path, 'w', compression=pyzipper.ZIP_DEFLATED, encryption=pyzipper.WZ_AES) as zf:
                zf.pwd = password.encode('utf-8')
                for file in input_files:
                    zf.write(file, arcname=os.path.basename(file))
        else:
            with pyzipper.AESZipFile(output_zip_path, 'w', compression=pyzipper.ZIP_DEFLATED, encryption=None) as zf:
                for file in input_files:
                    zf.write(file, arcname=os.path.basename(file))

    # Основная функция архивирования
    def do_archive():
        selected_mode = mode_var.get()
        directory = last_used_directory.get() or os.getcwd()
        
        try:
            password = None
            if use_password.get():
                password = entry_password.get().strip()
                if not password:
                    messagebox.showwarning("Ошибка", "Пароль не введён!")
                    return
            
            if selected_mode.lower() == "Каждый файл отдельно".lower():
                archive_each_file_separately(directory, password)
            elif selected_mode.lower() == "Все файлы в один архив".lower():
                archive_all_files_together(directory, password)
            else:
                raise ValueError("Режим архивации не выбран")
            
            messagebox.showinfo("Готово", "Архивация успешно выполнена.")
        except FileNotFoundError as fnfe:
            messagebox.showwarning("Предупреждение", str(fnfe))
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    # Интерфейс
    root = tk.Toplevel()
    root.title('Архиватор')
    root.geometry("250x300+100+150")
    root.iconbitmap(os.path.join(appdir,'osa.ico'))
    root.lift()
    root.attributes('-topmost', True)
    last_used_directory = StringVar(value=os.getcwd())  # Последняя используемая директория

    mode_var = StringVar(value="Каждый файл отдельно")  # Начальный режим по умолчанию
    use_password = BooleanVar(value=False)

    Label(root, text="Выберите режим архивации:").pack(pady=(10, 0))
    combo_modes = OptionMenu(root, mode_var, "Каждый файл отдельно", "Все файлы в один архив")#, command=lambda value: mode_var.set(value.lower()))  # Приводим строку к нижнему регистру
    combo_modes.pack(pady=(0, 10))

    Label(root, text="Защита паролем").pack()
    check_use_passwd = Checkbutton(root, text="Использовать пароль", variable=use_password,
                                command=lambda: toggle_password_field(use_password.get()))
    check_use_passwd.pack()

    label_password = Label(root, text="Введите пароль:", state=DISABLED)
    label_password.pack()

    entry_password = Entry(root, show='*', state=DISABLED)
    entry_password.pack()

    def toggle_password_field(active):
        if active:
            label_password.config(state=NORMAL)
            entry_password.config(state=NORMAL)
        else:
            label_password.config(state=DISABLED)
            entry_password.delete(0, END)
            entry_password.config(state=DISABLED)

    # Ярлык, отображающий текущую директорию
    Label(root, textvariable=last_used_directory).pack(pady=(10, 0))

    Button(root, text="Выбор директории", command=lambda: select_directory()).pack(pady=(10, 0))

    Button(root, text="Создать архив", command=do_archive).pack(pady=(10, 10))

    # Меняем директорию вручную
    def select_directory():
        chosen_dir = filedialog.askdirectory()
        if chosen_dir:
            last_used_directory.set(chosen_dir)

    root.mainloop()