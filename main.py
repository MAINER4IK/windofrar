import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
import subprocess
import platform
import os
import shutil
import sys

# Пути к файлам
slmgr_path = r"C:\Windows\System32\slmgr.vbs"
default_destination_dir = "C:/Program Files/WinRAR"

# Получаем абсолютный путь к скрипту или к временной папке, если это сборка с PyInstaller
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)

script_dir = resource_path("")
source_file = os.path.join(script_dir, "", "rarreg.key")
BAT_FILE_PATH = resource_path('activator.bat')  # Путь к activator.bat

# Функция для выполнения команд
def run_commands(commands):
    try:
        for command in commands:
            subprocess.run(
                command,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=subprocess.CREATE_NO_WINDOW,
                check=True
            )
        messagebox.showinfo("Успех", "Операция выполнена успешно!")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Ошибка", f"Ошибка при выполнении: {str(e)}")

# Функции активации/деактивации для Windows
def activate_windows():
    commands = [
        ["cscript", slmgr_path, "/ipk", "W269N-WFGWX-YVC9B-4J6C9-T83GX"],
        ["cscript", slmgr_path, "/skms", "kms8.msguides.com"],
        ["cscript", slmgr_path, "/ato"]
    ]
    run_commands(commands)

def deactivate_windows():
    commands = [
        ["cscript", slmgr_path, "/upk"],
        ["cscript", slmgr_path, "/cpky"]
    ]
    run_commands(commands)

# Функции активации/деактивации для WinRAR
def activate_winrar():
    try:
        if use_default.get():
            destination_dir = default_destination_dir
        else:
            destination_dir = path_entry.get()
            if not destination_dir:
                messagebox.showwarning("Предупреждение", "Вы должны указать путь к WinRAR.")
                return

        if os.path.isdir(destination_dir):
            destination_file = os.path.join(destination_dir, os.path.basename(source_file))
            if os.path.exists(destination_file):
                messagebox.showinfo("Информация", "WinRAR уже активирован.")
            else:
                shutil.copy(source_file, destination_dir)
                messagebox.showinfo("Успех", "WinRAR успешно активирован!")
        else:
            messagebox.showwarning("Предупреждение", "Указанный путь не является директорией.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка при активации: {e}")

def deactivate_winrar():
    try:
        if use_default.get():
            destination_dir = default_destination_dir
        else:
            destination_dir = path_entry.get()
            if not destination_dir:
                messagebox.showwarning("Предупреждение", "Вы должны указать путь к WinRAR.")
                return

        destination_file = os.path.join(destination_dir, os.path.basename(source_file))
        if os.path.exists(destination_file):
            os.remove(destination_file)
            messagebox.showinfo("Успех", "WinRAR успешно деактивирован!")
        else:
            messagebox.showwarning("Предупреждение", "Файл не найден для удаления.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка при деактивации WinRAR: {e}")

def check_activation_winrar():
    try:
        if use_default.get():
            destination_dir = default_destination_dir
        else:
            destination_dir = path_entry.get()
            if not destination_dir:
                messagebox.showwarning("Предупреждение", "Вы должны указать путь к WinRAR.")
                return

        destination_file = os.path.join(destination_dir, os.path.basename(source_file))
        if os.path.exists(destination_file):
            messagebox.showinfo("Статус", "WinRAR уже активирован.")
        else:
            messagebox.showinfo("Статус", "WinRAR не активирован.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка при проверке активации: {e}")

# Функции активации/деактивации для Office
def activate_office():
    try:
        subprocess.run([BAT_FILE_PATH], check=True, creationflags=subprocess.CREATE_NO_WINDOW)
        messagebox.showinfo("Статус", "Office активирован успешно.")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Ошибка", f"Не удалось запустить скрипт: {e}")

def deactivate_office():
    try:
        architecture = platform.architecture()[0]
        if architecture == '64bit':
            initial_command = 'cd %ProgramFiles%\\Microsoft Office\\Office16\\'
        elif architecture == '32bit':
            initial_command = 'cd %ProgramFiles(x86)%\\Microsoft Office\\Office16\\'
        else:
            messagebox.showerror("Ошибка", "Не удалось определить разрядность системы.")
            return

        # Команды для деактивации профессиональной плюс и стандартной версий
        deactivate_pro_plus_command = 'cscript ospp.vbs /unpkey:6F7TH'
        deactivate_standard_command = 'cscript ospp.vbs /unpkey:78VT3'
        
        final_command = f'{initial_command} && {deactivate_pro_plus_command} && {deactivate_standard_command}'

        subprocess.run(final_command, shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
        
        messagebox.showinfo("Статус", "Office деактивирован успешно.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось деактивировать Office: {e}")

# Главное меню с выбором активатора
def main_menu():
    main_window = tk.Tk()
    main_window.title("Выберите активатор")
    main_window.geometry("400x300")
    main_window.resizable(False, False)

    add_image(main_window, resource_path("images/windofrar.png"))

    windows_button = tk.Button(main_window, text="Активация Windows", command=lambda: open_windows_activator(main_window))
    windows_button.pack(pady=10)

    winrar_button = tk.Button(main_window, text="Активация WinRAR", command=lambda: open_winrar_activator(main_window))
    winrar_button.pack(pady=10)

    office_button = tk.Button(main_window, text="Активация Office", command=lambda: open_office_activator(main_window))
    office_button.pack(pady=10)

    main_window.mainloop()

# Окно активации Windows
def open_windows_activator(main_window):
    main_window.destroy()

    win_window = tk.Tk()
    win_window.title("Активатор Windows")
    win_window.geometry("400x300")
    win_window.resizable(False, False)

    add_image(win_window, resource_path("images/windows.png"))

    activate_win_button = tk.Button(win_window, text="Активировать Windows", command=activate_windows)
    activate_win_button.pack(pady=10)

    deactivate_win_button = tk.Button(win_window, text="Деактивировать Windows", command=deactivate_windows)
    deactivate_win_button.pack(pady=10)

    back_button = tk.Button(win_window, text="Назад", command=lambda: back_to_main_menu(win_window))
    back_button.pack(pady=10)

    win_window.mainloop()

# Окно активации WinRAR
def open_winrar_activator(main_window):
    global use_default, path_entry

    main_window.destroy()

    rar_window = tk.Tk()
    rar_window.title("Активатор WinRAR")
    rar_window.geometry("400x450")
    rar_window.resizable(False, False)

    add_image(rar_window, resource_path("images/winrar.png"))

    use_default = tk.BooleanVar(value=True)
    default_path_option = tk.Checkbutton(rar_window, text="Использовать путь по умолчанию для WinRAR", variable=use_default)
    default_path_option.pack(pady=5)

    path_label = tk.Label(rar_window, text="Или укажите свой путь:")
    path_label.pack()

    path_entry = tk.Entry(rar_window, width=50)
    path_entry.pack(pady=5)

    activate_rar_button = tk.Button(rar_window, text="Активировать WinRAR", command=activate_winrar)
    activate_rar_button.pack(pady=10)

    deactivate_rar_button = tk.Button(rar_window, text="Деактивировать WinRAR", command=deactivate_winrar)
    deactivate_rar_button.pack(pady=10)

    check_activation_button = tk.Button(rar_window, text="Проверить активацию", command=check_activation_winrar)
    check_activation_button.pack(pady=10)

    back_button = tk.Button(rar_window, text="Назад", command=lambda: back_to_main_menu(rar_window))
    back_button.pack(pady=10)

    rar_window.mainloop()

# Окно активации Office
def open_office_activator(main_window):
    main_window.destroy()

    office_window = tk.Tk()
    office_window.title("Активатор Office")
    office_window.geometry("400x300")
    office_window.resizable(False, False)

    add_image(office_window, resource_path("images/office.png"))

    activate_office_button = tk.Button(office_window, text="Активировать Office", command=activate_office)
    activate_office_button.pack(pady=10)

    deactivate_office_button = tk.Button(office_window, text="Деактивировать Office", command=deactivate_office)
    deactivate_office_button.pack(pady=10)

    back_button = tk.Button(office_window, text="Назад", command=lambda: back_to_main_menu(office_window))
    back_button.pack(pady=10)

    office_window.mainloop()

# Функция для возврата в главное меню
def back_to_main_menu(window):
    window.destroy()
    main_menu()

# Функция для добавления изображения
def add_image(window, image_path):
    try:
        image = Image.open(image_path)
        photo = ImageTk.PhotoImage(image)
        label = tk.Label(window, image=photo)
        label.image = photo
        label.pack()
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка при загрузке изображения: {e}")

if __name__ == "__main__":
    main_menu()
