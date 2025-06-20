import os
import shutil
from tkinter import *
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk
import pathlib

appdir = pathlib.Path(__file__).parent.resolve()

def move_selected_zips():
    global destination_dir
    
    # Проверяем наличие выбранного пути назначения
    if not destination_dir:
        messagebox.showwarning("Ошибка", "Сначала выберите директорию с архивами.")
        return

    # Получаем выбранные файлы
    selected_items = []
    for idx, var in enumerate(check_vars):
        if var.get():
            selected_items.append(zip_files[idx])

    if len(selected_items) == 0:
        messagebox.showwarning("Внимание", "Не выбрано ни одного файла для переноса.")
        return

    # Перемещаем выбранные архивы
    count_moved = 0
    for filename in selected_items:
        source_path = os.path.join(source_dir, filename)
        destination_path = os.path.join(destination_dir, filename)
        shutil.move(source_path, destination_path)
        count_moved += 1

    # Удаляем пустые папки, созданные заранее
    remove_empty_dirs()

    # Сообщаем пользователю количество перенесённых архивов
    messagebox.showinfo("Завершено", f"Было перемещено {count_moved} архивов.")

def toggle_all_selection():
    # Синхронизация состояния всех чекбоксов с состоянием "Выбрать все"
    selection_state = select_all_var.get()
    for var in check_vars:
        var.set(selection_state)

def create_archive_folder_if_needed():
    """
    Создает папку "Архив" только при наличии выбранных файлов и готовности к переносу.
    """
    global destination_dir
    counter = 0
    while True:
        candidate_dir = os.path.join(source_dir, f"Архив{counter}" if counter > 0 else "Архив")
        if not os.path.exists(candidate_dir):
            destination_dir = candidate_dir
            os.makedirs(destination_dir)
            break
        counter += 1

def remove_empty_dirs():
    """
    Убирает любые пустые папки "Архив*", если они были созданы, но не использованы.
    """
    for folder_name in os.listdir(source_dir):
        full_path = os.path.join(source_dir, folder_name)
        if os.path.isdir(full_path) and folder_name.startswith('Архив'):
            try:
                os.rmdir(full_path)  # Попытка удалить папку, если она пустая
            except OSError:
                pass  # Пропустить ошибку, если папка не пустая

def prepare_gui_for_move():
    global source_dir, destination_dir, zip_files, check_vars, select_all_var

    # Выбор директории с архивами
    source_dir = filedialog.askdirectory(title="Выберите директорию с архивами ZIP")
    if not source_dir:
        return  # Выход, если пользователь отменил выбор

    # Получаем список архивов
    zip_files = [f for f in os.listdir(source_dir) if f.endswith(".zip")]
    if len(zip_files) == 0:
        messagebox.showwarning("Внимание", "В директории нет архивов формата .zip.")
        return

    # Показываем диалог выбора файлов
    root = Toplevel()
    root.title("Выбор архивов для переноса")
    root.geometry("600x350+30+25")
    root.iconbitmap(os.path.join(appdir, 'osa.ico'))
    root.attributes('-topmost', True)

    # Основной фрейм для отображения файлов
    main_frame = ttk.Frame(root)
    main_frame.pack(fill=BOTH, expand=True)

    # Прокручиваемая область
    scrollbar = ttk.Scrollbar(main_frame)
    scrollbar.pack(side=RIGHT, fill=Y)

    # Канва для размещения чекбоксов
    container = Canvas(main_frame, yscrollcommand=scrollbar.set)
    container.pack(side=LEFT, fill=BOTH, expand=True)
    scrollbar.config(command=container.yview)

    # Рамка для чекбоксов
    content_frame = ttk.Frame(container)
    container.create_window((0, 0), window=content_frame, anchor=NW)

    # Глобальное перемещение чекбоксов
    select_all_var = IntVar()
    select_all_cb = ttk.Checkbutton(content_frame, text="Выбрать все", variable=select_all_var, command=toggle_all_selection)
    select_all_cb.grid(sticky=W)

    # Список чекбоксов для каждого файла
    check_vars = []
    for idx, filename in enumerate(zip_files):
        var = IntVar()
        check_vars.append(var)
        chk = ttk.Checkbutton(content_frame, text=filename, variable=var)
        chk.grid(sticky=W)

    # Скроллинг
    content_frame.update_idletasks()
    container.config(scrollregion=container.bbox(ALL))

    # Кнопка запуска переноса
    button = ttk.Button(root, text="Перенести выбранные архивы", command=lambda: [
        create_archive_folder_if_needed(),  # создаем папку только при нажатии кнопки
        move_selected_zips()
    ])
    button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    prepare_gui_for_move()# Запускаем подготовку и переход к выбору файлов


