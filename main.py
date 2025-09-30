import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
import subprocess, platform, os, shutil, sys

# Константы
SLMGR_PATH = r"C:\Windows\System32\slmgr.vbs"
DEFAULT_RAR_DIR = "C:/Program Files/WinRAR"
KMS_SERVER = "kms.digiboy.ir"

# Ключи активации Windows
WINDOWS_KEYS = {
    "Windows 11 Pro": "W269N-WFGWX-YVC9B-4J6C9-T83GX",
    "Windows 11 Home": "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99",
    "Windows 11 Enterprise": "NPPR9-FWDCX-D2C8J-H872K-2YT43",
    "Windows 10 Pro": "W269N-WFGWX-YVC9B-4J6C9-T83GX",
    "Windows 10 Home": "TX9XD-98N7V-6WMQ6-BX7FG-H8Q99",
    "Windows 10 Enterprise": "NPPR9-FWDCX-D2C8J-H872K-2YT43",
    "Windows 8.1 Pro": "GCRJD-8NW9H-F2CDX-CCM8D-9D6T9",
    "Windows 8.1 Enterprise": "MHF9N-XY6XB-WVXMC-BTDCT-MKKG7",
    "Windows 7 Pro": "FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4",
    "Windows 7 Enterprise": "33PXH-7Y6KF-2VJC9-XBBR8-HVTHH"
}

def resource_path(relative_path):
    try:
        return os.path.join(sys._MEIPASS, relative_path)
    except AttributeError:
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)

def run_slmgr_commands(commands):
    for command in commands:
        result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                              creationflags=subprocess.CREATE_NO_WINDOW, check=False)

        if result.returncode != 0:
            error_msg = result.stderr.decode('cp1251', errors='ignore') or f"Код ошибки: {result.returncode}"

            if result.returncode == 3221549093:  # STATUS_PIPE_BROKEN
                messagebox.showerror("Ошибка",
                    "Не удалось выполнить команду.\n\nВозможные причины:\n• Недостаточно прав\n• Антивирус блокирует\n• Поврежден slmgr.vbs\n\nПопробуйте:\n1. Запустить от имени администратора\n2. Отключить антивирус\n3. Перезагрузить компьютер")
            elif result.returncode == 3221225477:  # ERROR_ACCESS_DENIED
                messagebox.showerror("Ошибка доступа", "Недостаточно прав. Запустите от имени администратора.")
            else:
                messagebox.showerror("Ошибка", f"Ошибка выполнения:\n{error_msg}")
            return False

    messagebox.showinfo("Успех", "Операция выполнена успешно!")
    return True

def get_windows_info():
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion")
        edition, _ = winreg.QueryValueEx(key, "ProductName")
        winreg.CloseKey(key)
        return {"edition": edition, "release": platform.release()}
    except:
        return {"edition": "Неизвестная", "release": platform.release()}

def activate_windows_version(version_key):
    key = WINDOWS_KEYS.get(version_key)
    if not key:
        messagebox.showerror("Ошибка", "Ключ не найден")
        return

    commands = [["cscript", SLMGR_PATH, "/ipk", key],
                ["cscript", SLMGR_PATH, "/skms", KMS_SERVER],
                ["cscript", SLMGR_PATH, "/ato"]]
    run_slmgr_commands(commands)

def activate_current_windows():
    try:
        info = get_windows_info()
        edition = info.get("edition", "")

        # Определяем тип системы
        is_pro = "Pro" in edition or "Professional" in edition
        is_home = "Home" in edition
        is_enterprise = "Enterprise" in edition
        is_win11 = "11" in edition or info.get("release") == "11"

        if is_pro:
            version_key = "Windows 11 Pro" if is_win11 else "Windows 10 Pro"
        elif is_home:
            version_key = "Windows 11 Home" if is_win11 else "Windows 10 Home"
        elif is_enterprise:
            version_key = "Windows 11 Enterprise" if is_win11 else "Windows 10 Enterprise"
        else:
            messagebox.showinfo("Информация", f"Обнаружена редакция: {edition}\nИспользуйте ручной выбор.")
            return

        if version_key in WINDOWS_KEYS:
            activate_windows_version(version_key)
        else:
            messagebox.showwarning("Предупреждение", f"Ключ для '{edition}' не найден.")

    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка определения версии: {e}")

def deactivate_windows():
    commands = [["cscript", SLMGR_PATH, "/upk"], ["cscript", SLMGR_PATH, "/cpky"]]
    run_slmgr_commands(commands)

def get_rar_path():
    if use_default.get():
        return DEFAULT_RAR_DIR
    path = path_entry.get().strip()
    if not path:
        messagebox.showwarning("Предупреждение", "Укажите путь к WinRAR.")
        return None
    return path

def activate_winrar():
    try:
        destination_dir = get_rar_path()
        if not destination_dir: return

        if not os.path.isdir(destination_dir):
            messagebox.showwarning("Предупреждение", "Указанный путь не является директорией.")
            return

        source_file = resource_path("rarreg.key")
        destination_file = os.path.join(destination_dir, "rarreg.key")

        if os.path.exists(destination_file):
            messagebox.showinfo("Информация", "WinRAR уже активирован.")
        else:
            shutil.copy(source_file, destination_dir)
            messagebox.showinfo("Успех", "WinRAR успешно активирован!")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка активации: {e}")

def deactivate_winrar():
    try:
        destination_dir = get_rar_path()
        if not destination_dir: return

        destination_file = os.path.join(destination_dir, "rarreg.key")
        if os.path.exists(destination_file):
            os.remove(destination_file)
            messagebox.showinfo("Успех", "WinRAR успешно деактивирован!")
        else:
            messagebox.showwarning("Предупреждение", "Файл не найден.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка деактивации: {e}")

def check_winrar_activation():
    try:
        destination_dir = get_rar_path()
        if not destination_dir: return

        destination_file = os.path.join(destination_dir, "rarreg.key")
        status = "активирован" if os.path.exists(destination_file) else "не активирован"
        messagebox.showinfo("Статус", f"WinRAR {status}.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка проверки: {e}")

def activate_office():
    try:
        subprocess.run([resource_path('activator.bat')], check=True,
                      creationflags=subprocess.CREATE_NO_WINDOW)
        messagebox.showinfo("Статус", "Office активирован успешно.")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Ошибка", f"Не удалось запустить скрипт: {e}")

def deactivate_office():
    try:
        arch = platform.architecture()[0]
        office_path = '%ProgramFiles%\\Microsoft Office\\Office16\\' if arch == '64bit' else '%ProgramFiles(x86)%\\Microsoft Office\\Office16\\'

        if arch not in ['32bit', '64bit']:
            messagebox.showerror("Ошибка", "Не удалось определить разрядность системы.")
            return

        commands = [
            f'cd {office_path}',
            'cscript ospp.vbs /unpkey:6F7TH',
            'cscript ospp.vbs /unpkey:78VT3'
        ]

        subprocess.run(' && '.join(commands), shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
        messagebox.showinfo("Статус", "Office деактивирован успешно.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось деактивировать Office: {e}")

def create_window(title, image_name, buttons_config):
    window = tk.Tk()
    window.title(title)
    window.geometry("400x300")
    window.resizable(False, False)

    try:
        image = Image.open(resource_path(f"images/{image_name}"))
        photo = ImageTk.PhotoImage(image)
        label = tk.Label(window, image=photo)
        label.image = photo
        label.pack()
    except:
        pass

    # Создаем кнопки согласно конфигурации
    for btn_config in buttons_config:
        if len(btn_config) == 2:
            text, command = btn_config
            tk.Button(window, text=text, command=command).pack(pady=10)
        else:
            text, command, pady = btn_config
            tk.Button(window, text=text, command=command).pack(pady=pady)

    return window

def main_menu():
    def open_activator(activator_type, main_win):
        main_win.destroy()
        if activator_type == "windows":
            open_windows_activator()
        elif activator_type == "winrar":
            open_winrar_activator()
        elif activator_type == "office":
            open_office_activator()

    window = create_window("Выберите активатор", "windofrar.png", [
        ("Активация Windows", lambda: open_activator("windows", window)),
        ("Активация WinRAR", lambda: open_activator("winrar", window)),
        ("Активация Office", lambda: open_activator("office", window))
    ])
    window.mainloop()

def open_windows_activator():
    window = create_window("Активатор Windows", "windows.png", [
        ("Активировать Windows", activate_current_windows),
        ("Деактивировать Windows", deactivate_windows),
        ("Назад", lambda: back_to_main_menu(window))
    ])
    window.mainloop()

def open_winrar_activator():
    global use_default, path_entry

    window = tk.Tk()
    window.title("Активатор WinRAR")
    window.geometry("400x450")
    window.resizable(False, False)

    try:
        image = Image.open(resource_path("images/winrar.png"))
        photo = ImageTk.PhotoImage(image)
        label = tk.Label(window, image=photo)
        label.image = photo
        label.pack()
    except:
        pass

    use_default = tk.BooleanVar(value=True)
    tk.Checkbutton(window, text="Использовать путь по умолчанию для WinRAR", variable=use_default).pack(pady=5)

    tk.Label(window, text="Или укажите свой путь:").pack()
    path_entry = tk.Entry(window, width=50)
    path_entry.pack(pady=5)

    buttons = [
        ("Активировать WinRAR", activate_winrar, 10),
        ("Деактивировать WinRAR", deactivate_winrar, 10),
        ("Проверить активацию", check_winrar_activation, 10),
        ("Назад", lambda: back_to_main_menu(window), 10)
    ]

    for text, command, pady in buttons:
        tk.Button(window, text=text, command=command).pack(pady=pady)

    window.mainloop()

def open_office_activator():
    window = create_window("Активатор Office", "office.png", [
        ("Активировать Office", activate_office),
        ("Деактивировать Office", deactivate_office),
        ("Назад", lambda: back_to_main_menu(window))
    ])
    window.mainloop()

def back_to_main_menu(window):
    window.destroy()
    main_menu()

if __name__ == "__main__":
    main_menu()
