"""
Простая пересменка – GUI-программа для автоматизации подготовки к пересменке.
Версия: 5.1 (исправлены шаблоны путей, оптимизирован поиск)

Перед использованием:
1. Установите Python 3.9+ (с галочкой "Add Python to PATH").
2. Установите необходимые библиотеки:
   pip install -r requirements.txt
3. Скопируйте config.example.json в config.json и заполните актуальными данными.
4. Для сборки в exe:
   pyinstaller --onefile --windowed --name "Простая_пересменка" --icon=Peresmenca.ico --hidden-import=cv2 --collect-all opencv-python main.py
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import os
import sys
import datetime
import webbrowser
import subprocess
import time
import re
import random
import json
import winsound
import winreg
import numpy as np

# ------------------------------------------------------------------------------
#                 ЗНАЧЕНИЯ ПО УМОЛЧАНИЮ (если нет config.json)
# ------------------------------------------------------------------------------

DEFAULT_TELEMOST_URL = ""
DEFAULT_ENABLE_CAMERA = True
DEFAULT_ENABLE_DOCUMENTS = False
DEFAULT_NETWORK_BASE = r""
DEFAULT_RESERVE_BASE = r""
DEFAULT_DOKLAD_KEYWORD = ""
DEFAULT_SPRAVKA_KEYWORD = ""
DEFAULT_PREFERRED_MONITOR = 0
DEFAULT_DOCUMENT_SUBFOLDER = ""

# Тихий тест наушников (менять редко)
HEADPHONE_VOLUME = 0.2
HEADPHONE_FREQ = 800
HEADPHONE_DURATION = 0.15
HEADPHONE_INTERVAL = 0.5

# Размер окна камеры
CAMERA_WINDOW_WIDTH = 800
CAMERA_WINDOW_HEIGHT = 600

# ------------------------------------------------------------------------------
#           ПРОВЕРКА ДОСТУПНОСТИ ВНЕШНИХ БИБЛИОТЕК
# ------------------------------------------------------------------------------
try:
    import sounddevice as sd
    SOUNDDEVICE_AVAILABLE = True
except ImportError:
    SOUNDDEVICE_AVAILABLE = False
    sd = None

try:
    import cv2
    OPENCV_AVAILABLE = True
except ImportError:
    OPENCV_AVAILABLE = False
    cv2 = None

try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False
    win32com = None

# ------------------------------------------------------------------------------
#                         ЗАГРУЗКА КОНФИГУРАЦИИ
# ------------------------------------------------------------------------------
def load_config():
    """Загружает настройки из config.json, если он есть. Иначе – значения по умолчанию."""
    defaults = {
        "telemost_url": DEFAULT_TELEMOST_URL,
        "enable_camera_check": DEFAULT_ENABLE_CAMERA,
        "enable_documents": DEFAULT_ENABLE_DOCUMENTS,
        "network_base_path": DEFAULT_NETWORK_BASE,
        "reserve_base_path": DEFAULT_RESERVE_BASE,
        "doklad_keyword": DEFAULT_DOKLAD_KEYWORD,
        "spravka_keyword": DEFAULT_SPRAVKA_KEYWORD,
        "preferred_monitor": DEFAULT_PREFERRED_MONITOR,
        "document_subfolder": DEFAULT_DOCUMENT_SUBFOLDER,
        "main_path_template": "",
        "reserve_path_template": ""
    }

    # Ищем config.json рядом с exe или скриптом
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
    else:
        exe_dir = os.path.dirname(os.path.abspath(__file__))
    config_path = os.path.join(exe_dir, "config.json")

    if not os.path.exists(config_path):
        # Файла нет – работаем со значениями по умолчанию (всё выключено)
        return defaults

    try:
        with open(config_path, "r", encoding="utf-8-sig") as f:
            user_config = json.load(f)
    except json.JSONDecodeError as e:
        messagebox.showerror(
            "Ошибка config.json",
            f"Файл config.json повреждён и не может быть прочитан.\n\n"
            f"Ошибка: {e}\n\n"
            f"Программа продолжит работу с настройками по умолчанию (документы и Телемост будут отключены)."
        )
        return defaults
    except Exception as e:
        messagebox.showerror(
            "Ошибка config.json",
            f"Не удалось прочитать config.json.\n\n"
            f"Ошибка: {e}\n\n"
            f"Программа продолжит работу с настройками по умолчанию."
        )
        return defaults

    # Переопределяем только те ключи, которые есть в файле
    for key in defaults:
        if key in user_config:
            defaults[key] = user_config[key]

    # Проверяем критичные настройки и предупреждаем пользователя
    warnings = []
    if not defaults["telemost_url"]:
        warnings.append("• Не указана ссылка на Телемост (telemost_url)")
    if defaults["enable_documents"]:
        if not defaults["network_base_path"]:
            warnings.append("• Включены документы, но не указан основной сетевой путь (network_base_path)")
        if not defaults["doklad_keyword"] and not defaults["spravka_keyword"]:
            warnings.append("• Не указаны ключевые слова для поиска документов (doklad_keyword, spravka_keyword)")

    if warnings:
        messagebox.showwarning(
            "Предупреждение конфигурации",
            "Обнаружены возможные проблемы в настройках:\n\n" + "\n".join(warnings) +
            "\n\nПрограмма запустится, но некоторые функции могут не работать."
        )

    return defaults

# Применяем конфигурацию
config = load_config()
TELEMOST_URL = config["telemost_url"]
ENABLE_CAMERA_CHECK = config["enable_camera_check"]
ENABLE_DOCUMENTS = config["enable_documents"]
NETWORK_BASE_PATH = config["network_base_path"]
RESERVE_BASE_PATH = config["reserve_base_path"]
DOKLAD_KEYWORD = config["doklad_keyword"]
SPRAVKA_KEYWORD = config["spravka_keyword"]
PREFERRED_MONITOR = config["preferred_monitor"]
DOCUMENT_SUBFOLDER = config["document_subfolder"]
MAIN_PATH_TEMPLATE = config["main_path_template"]
RESERVE_PATH_TEMPLATE = config["reserve_path_template"]

# ------------------------------------------------------------------------------
#                       ОСНОВНОЙ КЛАСС ПРИЛОЖЕНИЯ
# ------------------------------------------------------------------------------
class ShiftChangeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Простая пересменка")
        self.root.geometry("620x400")
        self.root.resizable(False, False)
        self.root.configure(bg='#f0f0f0')

        # Состояния тестов
        self.headphone_test_active = False
        self.mic_test_active = False
        self.camera_test_active = False

        self.headphone_thread = None
        self.stop_headphone_flag = threading.Event()

        self.mic_stream = None
        self.stop_mic_flag = threading.Event()

        self.camera_thread = None
        self.stop_camera_flag = threading.Event()

        self.camera_window_closed = False
        self.camera_thread = None

        self.show_initial_screen()

    # --------------------------------------------------------------------------
    #                         УПРАВЛЕНИЕ ОКНАМИ
    # --------------------------------------------------------------------------
    def clear_window(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def show_initial_screen(self):
        self.clear_window()
        self.root.title("Подготовка к пересменке")

        main_frame = ttk.Frame(self.root, padding=30)
        main_frame.pack(expand=True, fill=tk.BOTH)

        label = ttk.Label(
            main_frame,
            text="Подключите наушники с микрофоном\nи нажмите «Продолжить»",
            font=("Segoe UI", 14),
            wraplength=500,
            justify=tk.CENTER
        )
        label.pack(pady=(40, 30))

        self.btn_continue = ttk.Button(
            main_frame,
            text="Продолжить",
            command=self.check_devices_and_proceed,
            style="Accent.TButton"
        )
        self.btn_continue.pack(pady=20)

        style = ttk.Style()
        style.configure("Accent.TButton", font=("Segoe UI", 12, "bold"), padding=10)
        self.root.after(100, self._bring_to_front)

    def _bring_to_front(self):
        """Поднимает окно на передний план и подаёт системный звук уведомления."""
        self.root.attributes('-topmost', True)
        self.root.lift()
        self.root.focus_force()
        self.root.attributes('-topmost', False)
        winsound.MessageBeep(winsound.MB_ICONINFORMATION)

    def check_devices_and_proceed(self):
        self.show_main_menu()

    def show_main_menu(self):
        self.clear_window()
        self.root.title("Главное меню – Проверка оборудования")

        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(expand=True, fill=tk.BOTH)

        ttk.Label(
            main_frame,
            text="Проверьте оборудование перед пересменкой",
            font=("Segoe UI", 11, "bold"),
            anchor=tk.CENTER
        ).pack(pady=(0, 20))

        tests_frame = ttk.Frame(main_frame)
        tests_frame.pack(pady=10)

        self.btn_headphones = ttk.Button(
            tests_frame,
            text="🎧 Проверить наушники",
            command=self.toggle_headphone_test,
            width=22
        )
        self.btn_headphones.grid(row=0, column=0, padx=10, pady=5)

        self.btn_microphone = ttk.Button(
            tests_frame,
            text="🎤 Проверить микрофон",
            command=self.toggle_microphone_test,
            width=22
        )
        self.btn_microphone.grid(row=0, column=1, padx=10, pady=5)

        if ENABLE_CAMERA_CHECK and OPENCV_AVAILABLE:
            self.btn_camera = ttk.Button(
                tests_frame,
                text="📷 Проверить камеру",
                command=self.toggle_camera_test,
                width=22
            )
            self.btn_camera.grid(row=0, column=2, padx=10, pady=5)
        elif ENABLE_CAMERA_CHECK:
            btn_disabled = ttk.Button(
                tests_frame,
                text="📷 Камера (нет OpenCV)",
                state=tk.DISABLED,
                width=22
            )
            btn_disabled.grid(row=0, column=2, padx=10, pady=5)

        self.btn_start = ttk.Button(
            main_frame,
            text="▶  Начать пересменку",
            command=self.start_shift_sequence,
            style="Accent.TButton"
        )
        self.btn_start.pack(pady=30)

        style = ttk.Style()
        style.configure("Accent.TButton", font=("Segoe UI", 12, "bold"), padding=10)

    def show_reminder(self):
        self.clear_window()
        self.root.title("Внимание!")

        frame = ttk.Frame(self.root, padding=30)
        frame.pack(expand=True, fill=tk.BOTH)

        reminder_text = (
            "Сейчас откроется Телемост.\n"
            "Пожалуйста, разрешите доступ к микрофону и камере при необходимости.\n"
            "Не забудьте ввести наименование вашего отдела/должности."
        )

        ttk.Label(
            frame,
            text=reminder_text,
            font=("Segoe UI", 12),
            justify=tk.CENTER,
            wraplength=500
        ).pack(pady=(0, 30))

        ttk.Button(
            frame,
            text="Продолжить",
            command=self.open_browser_and_docs,
            style="Accent.TButton"
        ).pack()

    def _find_telemost_exe(self):
        """Ищет исполняемый файл Яндекс.Телемоста через реестр и стандартные папки."""
        # Способ 1: ищем в реестре команду для протокола telemost://
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"telemost\shell\open\command") as key:
                command = winreg.QueryValue(key, "")
                if command:
                    exe_path = command.split('"')[1]
                    if os.path.exists(exe_path):
                        return exe_path
        except Exception:
            pass

        # Способ 2: стандартные пути
        paths = [
            os.path.join(os.environ.get("LOCALAPPDATA", ""),
                         "Programs", "YandexTelemost", "Telemost.exe"),
            os.path.join(os.environ.get("PROGRAMFILES", "C:\\Program Files"),
                         "Yandex", "YandexTelemost", "Telemost.exe"),
            os.path.join(os.environ.get("PROGRAMFILES(X86)", "C:\\Program Files (x86)"),
                         "Yandex", "YandexTelemost", "Telemost.exe"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""),
                         "Yandex", "YandexTelemost", "Application", "Telemost.exe"),
        ]
        for p in paths:
            if os.path.exists(p):
                return p
        return None

    def open_browser_and_docs(self):
        try:
            if ENABLE_DOCUMENTS:
                if not NETWORK_BASE_PATH:
                    messagebox.showwarning(
                        "Документы отключены",
                        "В config.json включено открытие документов, но не указан основной сетевой путь.\n"
                        "Документы не будут открыты."
                    )
                else:
                    self.open_word_documents()

            self.open_telemost(TELEMOST_URL)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{e}")
        finally:
            self.root.destroy()
            sys.exit(0)

    # --------------------------------------------------------------------------
    #                      ЛОГИКА ТЕСТОВ ОБОРУДОВАНИЯ
    # --------------------------------------------------------------------------

    # ---------------------------- Тихий тест наушников -------------------------
    def toggle_headphone_test(self):
        if not self.headphone_test_active:
            if not SOUNDDEVICE_AVAILABLE:
                messagebox.showerror("Ошибка", "Библиотека sounddevice не установлена.\nУстановите: pip install sounddevice")
                return
            try:
                self.headphone_test_active = True
                self.stop_headphone_flag.clear()
                self.btn_headphones.config(text="🛑 Остановить тест")
                self.headphone_thread = threading.Thread(
                    target=self._play_tone_loop, daemon=True
                )
                self.headphone_thread.start()
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось запустить тест наушников:\n{e}")
                self.headphone_test_active = False
                self.btn_headphones.config(text="🎧 Проверить наушники")
        else:
            self.stop_headphone_flag.set()
            if self.headphone_thread and self.headphone_thread.is_alive():
                self.headphone_thread.join(timeout=1.0)
            self.headphone_test_active = False
            self.btn_headphones.config(text="🎧 Проверить наушники")

    def _play_tone_loop(self):
        sample_rate = 22050
        t = np.linspace(0, HEADPHONE_DURATION, int(sample_rate * HEADPHONE_DURATION), endpoint=False)
        tone = (HEADPHONE_VOLUME * np.sin(2 * np.pi * HEADPHONE_FREQ * t)).astype(np.float32)

        while not self.stop_headphone_flag.is_set():
            try:
                sd.play(tone, samplerate=sample_rate, blocking=False)
                time.sleep(HEADPHONE_DURATION + HEADPHONE_INTERVAL)
            except:
                break
        sd.stop()

    # ---------------------------- Микрофон -------------------------------------
    def toggle_microphone_test(self):
        if not self.mic_test_active:
            if not SOUNDDEVICE_AVAILABLE:
                messagebox.showerror("Ошибка", "Библиотека sounddevice не установлена.\nУстановите: pip install sounddevice")
                return
            try:
                sd._terminate()
                sd._initialize()
                self.mic_test_active = True
                self.stop_mic_flag.clear()
                self.btn_microphone.config(text="🛑 Остановить тест")
                self.mic_stream = sd.Stream(
                    samplerate=44100,
                    channels=1,
                    dtype='int16',
                    callback=self._mic_callback
                )
                self.mic_stream.start()
            except sd.PortAudioError as e:
                if "querying device" in str(e).lower() or "invalid device" in str(e).lower():
                    messagebox.showwarning(
                        "Микрофон недоступен",
                        "Не удалось найти микрофон.\nПроверьте, подключена ли гарнитура и не занята ли она другим приложением."
                    )
                else:
                    messagebox.showerror("Ошибка", f"Не удалось запустить тест микрофона:\n{e}")
                self.mic_test_active = False
                self.btn_microphone.config(text="🎤 Проверить микрофон")
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось запустить тест микрофона:\n{e}")
                self.mic_test_active = False
                self.btn_microphone.config(text="🎤 Проверить микрофон")
        else:
            self.stop_mic_flag.set()
            if self.mic_stream:
                try:
                    self.mic_stream.stop()
                    self.mic_stream.close()
                except:
                    pass
                self.mic_stream = None
            self.mic_test_active = False
            self.btn_microphone.config(text="🎤 Проверить микрофон")

    def _mic_callback(self, indata, outdata, frames, time_info, status):
        if self.stop_mic_flag.is_set():
            raise sd.CallbackStop
        outdata[:] = indata

    # ---------------------------- Камера ---------------------------------------
    def toggle_camera_test(self):
        if not self.camera_test_active:
            if not OPENCV_AVAILABLE:
                messagebox.showerror("Ошибка", "OpenCV не установлен.\nУстановите: pip install opencv-python")
                return
            try:
                self.camera_test_active = True
                self.stop_camera_flag.clear()
                self.btn_camera.config(text="🛑 Остановить тест")
                self.camera_thread = threading.Thread(
                    target=self._camera_preview, daemon=True
                )
                self.camera_thread.start()
                self.monitor_camera_thread()
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось запустить камеру:\n{e}")
                self._finish_camera_test()
        else:
            self.stop_camera_flag.set()
            self._finish_camera_test()

    def _camera_preview(self):
        import ctypes

        cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
        if not cap.isOpened():
            self.root.after(0, lambda: messagebox.showerror("Ошибка", "Не удалось открыть веб-камеру."))
            self.root.after(0, self._finish_camera_test)
            return

        temp_name = "CameraWindow"
        win_name = "Тест камеры (закройте окно или нажмите 'Остановить')"
        cv2.namedWindow(temp_name, cv2.WINDOW_NORMAL)
        cv2.resizeWindow(temp_name, CAMERA_WINDOW_WIDTH, CAMERA_WINDOW_HEIGHT)

        hwnd = ctypes.windll.user32.FindWindowW(None, temp_name)
        if hwnd:
            ctypes.windll.user32.SetWindowTextW(hwnd, win_name)

        while not self.stop_camera_flag.is_set():
            ret, frame = cap.read()
            if not ret:
                break
            if cv2.getWindowProperty(temp_name, cv2.WND_PROP_VISIBLE) < 1:
                break
            cv2.imshow(temp_name, frame)
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break

        cap.release()
        cv2.destroyWindow(temp_name)

    def monitor_camera_thread(self):
        if self.camera_thread is not None and not self.camera_thread.is_alive():
            if not self.stop_camera_flag.is_set():
                self.root.after(0, self._finish_camera_test)
        else:
            if self.camera_test_active:
                self.root.after(200, self.monitor_camera_thread)

    def _finish_camera_test(self):
        self.camera_test_active = False
        self.stop_camera_flag.set()
        self.btn_camera.config(text="📷 Проверить камеру")

    # --------------------------------------------------------------------------
    #                   ПЕРЕХОД К ПЕРЕСМЕНКЕ
    # --------------------------------------------------------------------------
    def start_shift_sequence(self):
        if self.headphone_test_active:
            self.stop_headphone_flag.set()
        if self.mic_test_active:
            self.stop_mic_flag.set()
            if self.mic_stream:
                try:
                    self.mic_stream.stop()
                    self.mic_stream.close()
                except:
                    pass
                self.mic_stream = None
        if self.camera_test_active:
            self.stop_camera_flag.set()

        self.headphone_test_active = False
        self.mic_test_active = False
        self.camera_test_active = False

        time.sleep(0.2)
        self.show_reminder()

    def open_word_documents(self):
        today = datetime.datetime.now()
        # Если сейчас до полудня – ищем документы за вчерашний день
        if today.hour < 12:
            search_date = today - datetime.timedelta(days=1)
        else:
            search_date = today

        year = search_date.strftime("%Y")
        month_num = search_date.strftime("%m")
        month_names = [
            "январь", "февраль", "март", "апрель", "май", "июнь",
            "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"
        ]
        month_name = month_names[int(month_num) - 1]
        month_folder = f"{month_num} {month_name}"
        date_short = today.strftime("%d.%m.%y")
        date_full = today.strftime("%d.%m.%Y")
        search_date_short = search_date.strftime("%d.%m.%y")

        # Формируем пути с помощью шаблонов (если заданы), иначе – старым способом
        def build_path(template, default_path_func):
            if template:
                path = template
                replacements = {
                    "{network_base}": NETWORK_BASE_PATH,
                    "{reserve_base}": RESERVE_BASE_PATH,
                    "{year}": year,
                    "{month}": month_num,
                    "{month_name}": month_name,
                    "{date_short}": date_short,
                    "{date_full}": date_full,
                    "{search_date_short}": search_date_short,
                }
                for key, value in replacements.items():
                    path = path.replace(key, value)
                return path
            else:
                return default_path_func()

        # Основной путь
        if MAIN_PATH_TEMPLATE:
            main_dir = build_path(MAIN_PATH_TEMPLATE, lambda: os.path.join(NETWORK_BASE_PATH, year, month_folder, search_date_short, DOCUMENT_SUBFOLDER))
        else:
            main_dir = os.path.join(NETWORK_BASE_PATH, year, month_folder, search_date_short, DOCUMENT_SUBFOLDER)

        # Резервный путь
        if RESERVE_PATH_TEMPLATE:
            reserve_dir = build_path(RESERVE_PATH_TEMPLATE, lambda: os.path.join(RESERVE_BASE_PATH, today.strftime("%Y"), month_folder))
        else:
            reserve_dir = os.path.join(RESERVE_BASE_PATH, today.strftime("%Y"), month_folder)

        doklad_file = None
        spravka_file = None

        # Проверяем существование папок
        if not os.path.exists(main_dir) and not (RESERVE_PATH_TEMPLATE or os.path.exists(reserve_dir)):
            messagebox.showwarning(
                "Папка не найдена",
                f"Не удалось найти папку с документами.\n\n"
                f"Ожидаемый путь:\n{main_dir}\n\n"
                f"Проверьте настройки сетевого пути и шаблона в config.json.\n"
                f"Если вы используете шаблоны, убедитесь, что они содержат корректные переменные."
            )
            return

        # Поиск доклада в основной папке
        if os.path.exists(main_dir):
            try:
                for f in os.listdir(main_dir):
                    if f.startswith(DOKLAD_KEYWORD):
                        doklad_file = os.path.join(main_dir, f)
                        break
            except Exception as e:
                messagebox.showwarning("Поиск доклада", f"Ошибка: {e}")

        # Поиск справки в основной папке
        if os.path.exists(main_dir):
            try:
                for f in os.listdir(main_dir):
                    if SPRAVKA_KEYWORD in f:
                        spravka_file = os.path.join(main_dir, f)
                        break
            except Exception as e:
                messagebox.showwarning("Поиск справки", f"Ошибка при поиске справки: {e}")

        # Если справка не найдена – ищем в резервной папке с датой в имени
        if not spravka_file and os.path.exists(reserve_dir):
            try:
                for f in os.listdir(reserve_dir):
                    if SPRAVKA_KEYWORD in f and date_full in f:
                        spravka_file = os.path.join(reserve_dir, f)
                        break
            except Exception as e:
                messagebox.showwarning("Поиск справки (резерв)", f"Ошибка: {e}")

        if not doklad_file and not spravka_file:
            messagebox.showwarning("Документы не найдены", "Не удалось найти ни доклад, ни справку.")
            return

        # Открываем файлы в Word
        files_to_open = []
        if doklad_file:
            files_to_open.append(("left", doklad_file))
        if spravka_file:
            files_to_open.append(("right", spravka_file))

        for side, path in files_to_open:
            try:
                if WIN32COM_AVAILABLE:
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = True
                    try:
                        word.Options.DisableProtectedViewWarning = True
                    except:
                        pass
                    word.Documents.Open(path)
                    self._position_word_window(word, side)
                else:
                    os.startfile(path)
            except Exception as e:
                messagebox.showerror("Ошибка открытия", f"Не удалось открыть {path}:\n{e}")

    def _position_word_window(self, word, side):
        """Размещает окно Word на левой или правой половине выбранного монитора."""
        try:
            import win32gui
            import win32api
            import win32con
            import ctypes
            from ctypes import wintypes

            time.sleep(1.5)
            hwnd = word.ActiveWindow.Hwnd

            monitors = []
            def monitor_enum_callback(hMonitor, hdcMonitor, lprcMonitor, dwData):
                mi = win32api.GetMonitorInfo(hMonitor)
                monitors.append(mi)
                return True

            MonitorEnumProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p,
                                                  ctypes.POINTER(wintypes.RECT), ctypes.c_void_p)
            callback = MonitorEnumProc(monitor_enum_callback)
            ctypes.windll.user32.EnumDisplayMonitors(None, None, callback, 0)

            monitors.sort(key=lambda m: (m['Work'][0], m['Work'][1]))

            monitor_index = PREFERRED_MONITOR
            if monitor_index < 0 or monitor_index >= len(monitors):
                monitor_index = 0
            monitor_info = monitors[monitor_index]

            work_area = monitor_info['Work']
            screen_left = work_area[0]
            screen_top = work_area[1]
            screen_width = work_area[2] - work_area[0]
            screen_height = work_area[3] - work_area[1]

            half_width = screen_width // 2

            if side == "left":
                left = screen_left
            else:
                left = screen_left + half_width

            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.MoveWindow(hwnd, left, screen_top, half_width, screen_height, True)

        except Exception:
            pass

    # ---------------------------- Браузер / Телемост ---------------------------
    def _extract_meeting_id(self, url):
        match = re.search(r'telemost\.yandex\.ru/(?:j/)?(\d+)', url)
        if match:
            return match.group(1)
        return None

    def _generate_session_id(self):
        part1 = random.randint(10**15, 10**16-1)
        part2 = random.randint(10**16, 10**17-1)
        return f"{part1}-{part2}"

    def open_telemost(self, url):

        if not url:
            messagebox.showwarning(
                "Ссылка не задана",
                "В config.json не указана ссылка на Телемост.\n"
                "Браузер не будет открыт."
            )
            return
        
        meeting_id = self._extract_meeting_id(url)
        if not meeting_id:
            self._fallback_browser_open(url)
            return

        telemost_exe = self._find_telemost_exe()
        if telemost_exe:
            session_id = self._generate_session_id()
            conf_url = f"telemost://https//telemost.yandex.ru/j/{meeting_id}?sessionid={session_id}"
            try:
                subprocess.Popen([telemost_exe, f"--conf-url={conf_url}"])
                return
            except Exception:
                pass

        self._fallback_browser_open(url)

    def _fallback_browser_open(self, url):
        yandex_paths = [
            os.path.join(os.environ.get("LOCALAPPDATA", ""),
                         "Yandex", "YandexBrowser", "Application", "browser.exe"),
            os.path.join(os.environ.get("PROGRAMFILES", "C:\\Program Files"),
                         "Yandex", "YandexBrowser", "Application", "browser.exe"),
            os.path.join(os.environ.get("PROGRAMFILES(X86)", "C:\\Program Files (x86)"),
                         "Yandex", "YandexBrowser", "Application", "browser.exe"),
        ]
        for path in yandex_paths:
            if os.path.exists(path):
                try:
                    subprocess.Popen([path, "--incognito", url])
                    return
                except Exception:
                    pass
        webbrowser.open(url)

# ------------------------------------------------------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = ShiftChangeApp(root)
    root.mainloop()