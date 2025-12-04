import sys
import os
import json
import time
import logging
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Tuple, Optional
import threading
import queue
import base64
from io import BytesIO

# Для работы с Excel
import pandas as pd
import openpyxl

# Для автоматизации интерфейса
import pyautogui
import keyboard

# Для парсинга дат
from dateutil.parser import parse as date_parse

# Для интерфейса
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# Для работы с изображениями
from PIL import Image, ImageGrab

# Для OCR
import pytesseract

# Для работы с буфером обмена
try:
    import pyperclip

    HAS_PYPERCLIP = True
except ImportError:
    HAS_PYPERCLIP = False
    logging.warning(
        "Модуль pyperclip не установлен. Установите его для лучшей работы с буфером обмена: pip install pyperclip")

pyautogui.FAILSAFE = True


# ================== КОНФИГУРАЦИЯ ==================
class Config:
    def __init__(self):
        self.actions_file = "form_actions.json"
        self.excel_file = ""
        self.start_row = 0
        self.speed_factor = 1.0
        self.log_level = "INFO"
        self.use_image_recognition = True
        self.verify_input = True  # Новая опция: проверять введенные данные
        self.max_attempts = 5  # Увеличено для надежности
        self.ocr_lang = 'rus'  # Язык для OCR
        self.use_clipboard = True  # Использовать буфер обмена для вставки

    def save(self, filename: str = "config.json"):
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(self.__dict__, f, indent=2, ensure_ascii=False)

    @classmethod
    def load(cls, filename: str = "config.json"):
        if os.path.exists(filename):
            with open(filename, 'r', encoding='utf-8') as f:
                data = json.load(f)
            config = cls()
            config.__dict__.update(data)
            return config
        return cls()


# ================== ТИПЫ ДАННЫХ ==================
class FieldType:
    LAST_NAME = "Фамилия"
    FIRST_NAME = "Имя"
    MIDDLE_NAME = "Отчество"
    BIRTH_DAY = "Дата рождения (день)"
    BIRTH_MONTH = "Дата рождения (месяц)"
    BIRTH_YEAR = "Дата рождения (год)"


class FormField:
    def __init__(self, name: str, field_type: str, screen_position: Tuple[int, int], size: Tuple[int, int] = (300, 50),
                 image_data: Optional[str] = None, click_offset: Tuple[int, int] = (10, 10)):
        self.name = name
        self.field_type = field_type
        self.screen_position = screen_position
        self.size = size
        self.image_data = image_data
        self.click_offset = click_offset

    def get_click_position(self) -> Tuple[int, int]:
        x, y = self.screen_position
        w, h = self.size
        center_x, center_y = x + w // 2, y + h // 2
        return (center_x + self.click_offset[0], center_y + self.click_offset[1])

    def to_dict(self):
        return {
            'name': self.name,
            'field_type': self.field_type,
            'screen_position': self.screen_position,
            'size': self.size,
            'image_data': self.image_data,
            'click_offset': self.click_offset
        }

    @classmethod
    def from_dict(cls, data: dict):
        return cls(
            name=data['name'],
            field_type=data['field_type'],
            screen_position=tuple(data['screen_position']),
            size=tuple(data['size']),
            image_data=data.get('image_data'),
            click_offset=tuple(data.get('click_offset', (10, 10)))
        )


class FormAction:
    def __init__(self, field: FormField, value: str, delay_before: float = 0.5, delay_after: float = 0.5):
        self.field = field
        self.value = str(value) if value is not None else ""
        self.delay_before = delay_before
        self.delay_after = delay_after

    def get_field_region(self) -> Tuple[int, int, int, int]:
        field_x, field_y = self.field.screen_position
        field_w, field_h = self.field.size
        return (field_x, field_y, field_x + field_w, field_y + field_h)

    def read_field_content(self, region: Tuple[int, int, int, int], lang: str = 'rus') -> str:
        """Чтение содержимого поля с помощью OCR"""
        try:
            screenshot = ImageGrab.grab(bbox=region)
            text = pytesseract.image_to_string(screenshot, lang=lang).strip()
            return text
        except Exception as e:
            logging.warning(f"Ошибка OCR: {e}")
            return ""

    def verify_field_content(self, expected_value: str, region: Tuple[int, int, int, int], lang: str = 'rus') -> bool:
        """Проверить содержимое поля с помощью OCR"""
        actual_value = self.read_field_content(region, lang)
        if actual_value == expected_value:
            return True
        else:
            logging.debug(f"Проверка не пройдена. Ожидалось: '{expected_value}', получено: '{actual_value}'")
            return False

    def clear_field(self, region: Tuple[int, int, int, int], speed_factor: float = 1.0, ocr_lang: str = 'rus') -> bool:
        """Очистить поле ввода"""
        # Попробовать разные методы очистки
        methods = [
            lambda: pyautogui.hotkey('ctrl', 'a'),  # Ctrl+A для выделения всего
            lambda: pyautogui.click(),
            lambda: pyautogui.doubleClick(),
            lambda: pyautogui.tripleClick(),  # Для полей, где тройной клик выделяет всё
        ]

        for method in methods:
            try:
                method()
                time.sleep(0.1 * speed_factor)
                pyautogui.press('delete')
                time.sleep(0.2 * speed_factor)
                pyautogui.press('backspace', presses=3)  # На всякий случай
                time.sleep(0.3 * speed_factor)

                # Проверить, очистилось ли
                content = self.read_field_content(region, ocr_lang)
                if not content.strip():
                    return True
            except Exception as e:
                logging.debug(f"Метод очистки не сработал: {e}")
                continue

        # Если стандартные методы не сработали, попробовать более агрессивный подход
        try:
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.1)
            for _ in range(10):  # Много раз нажать delete/backspace
                pyautogui.press('delete')
                pyautogui.press('backspace')
                time.sleep(0.05)

            content = self.read_field_content(region, ocr_lang)
            if not content.strip():
                return True
        except Exception as e:
            logging.warning(f"Не удалось очистить поле: {e}")

        return False

    def paste_with_clipboard(self, text: str, speed_factor: float = 1.0) -> bool:
        """Вставить текст используя буфер обмена"""
        if not HAS_PYPERCLIP:
            logging.warning("pyperclip не установлен, использую typewrite")
            pyautogui.typewrite(text, interval=0.05 * speed_factor)
            return True

        try:
            # Сохранить текущий буфер обмена
            original_clipboard = ""
            try:
                original_clipboard = pyperclip.paste()
            except:
                pass

            # Копировать текст в буфер обмена
            pyperclip.copy(text)
            time.sleep(0.1 * speed_factor)  # Дать время для копирования

            # Вставить из буфера обмена
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.2 * speed_factor)  # Дать время для вставки

            # Восстановить оригинальный буфер обмена
            if original_clipboard:
                time.sleep(0.1)
                pyperclip.copy(original_clipboard)

            return True
        except Exception as e:
            logging.error(f"Ошибка при работе с буфером обмена: {e}")
            # Fallback: использовать обычный ввод
            pyautogui.typewrite(text, interval=0.05 * speed_factor)
            return False

    def execute(self, speed_factor: float = 1.0, use_image: bool = False, verify: bool = False,
                max_attempts: int = 5, ocr_lang: str = 'rus', use_clipboard: bool = True) -> bool:
        """Выполнить действие с проверками"""
        for attempt in range(max_attempts):
            try:
                time.sleep(self.delay_before * speed_factor)

                if use_image and self.field.image_data:
                    img_bytes = base64.b64decode(self.field.image_data)
                    img = Image.open(BytesIO(img_bytes))
                    location = None
                    for confidence in [0.9, 0.8, 0.7]:
                        location = pyautogui.locateOnScreen(img, confidence=confidence, grayscale=True)
                        if location:
                            break
                    if location:
                        center_x, center_y = pyautogui.center(location)
                        click_x = center_x + self.field.click_offset[0]
                        click_y = center_y + self.field.click_offset[1]
                    else:
                        logging.warning(f"Изображение не найдено для поля {self.field.name}, использую координаты")
                        click_x, click_y = self.field.get_click_position()
                else:
                    click_x, click_y = self.field.get_click_position()

                pyautogui.moveTo(click_x, click_y, duration=0.3 * speed_factor)
                time.sleep(0.2 * speed_factor)

                # Сначала кликнуть, чтобы активировать поле
                pyautogui.click()
                time.sleep(0.1 * speed_factor)

                # Потом тройной клик для выделения всего текста в поле
                pyautogui.tripleClick()
                time.sleep(0.2 * speed_factor)

                # Очистка поля
                region = self.get_field_region()
                if not self.clear_field(region, speed_factor, ocr_lang):
                    logging.warning(f"Не удалось очистить поле {self.field.name}")
                    continue

                # Ввод текста
                if self.value:
                    if use_clipboard:
                        self.paste_with_clipboard(self.value, speed_factor)
                    else:
                        pyautogui.typewrite(self.value, interval=0.05 * speed_factor)

                time.sleep(0.3 * speed_factor)  # Время на обновление поля

                if verify:
                    time.sleep(0.5 * speed_factor)  # Дополнительная задержка для обновления
                    if self.verify_field_content(self.value, region, ocr_lang):
                        logging.info(f"✓ Поле '{self.field.name}' успешно заполнено значением '{self.value}'")
                        time.sleep(self.delay_after * speed_factor)
                        return True
                    else:
                        logging.warning(f"Попытка {attempt + 1}/{max_attempts} не удалась для поля '{self.field.name}'")
                        if attempt < max_attempts - 1:
                            time.sleep(1.0)  # Пауза перед повтором
                            continue
                else:
                    time.sleep(self.delay_after * speed_factor)
                    return True

            except Exception as e:
                logging.error(f"Ошибка при заполнении поля '{self.field.name}': {e}")
                if attempt < max_attempts - 1:
                    time.sleep(2.0)  # Дольше ждем
                    continue

        logging.error(f"Не удалось заполнить поле '{self.field.name}' после {max_attempts} попыток")
        return False


# ================== МЕНЕДЖЕР ФОРМ ==================
class FormManager:
    def __init__(self):
        self.fields: List[FormField] = []
        self.is_recording = False
        self.record_start_time = 0

    def start_recording(self, use_image: bool = False):
        self.is_recording = True
        self.fields = []
        self.record_start_time = time.time()
        self.use_image = use_image
        logging.info("Запись начата. Используйте горячие клавиши для записи полей.")

    def stop_recording(self):
        self.is_recording = False

    def record_field(self, field_type: str, position: Tuple[int, int]):
        image_data = None
        if hasattr(self, 'use_image') and self.use_image:
            x, y = position
            w, h = 200, 60
            screenshot = pyautogui.screenshot(region=(x - w // 2, y - h // 2, w, h))
            buffered = BytesIO()
            screenshot.save(buffered, format="PNG")
            image_data = base64.b64encode(buffered.getvalue()).decode('utf-8')
        field = FormField(
            name=field_type,
            field_type=field_type,
            screen_position=(position[0] - 100, position[1] - 15),
            size=(200, 30),
            image_data=image_data,
            click_offset=(0, 0)
        )
        self.fields.append(field)
        logging.info(f"Записано поле: {field_type} на позиции {position}")

    def save_fields(self, filename: str):
        data = {
            'fields': [field.to_dict() for field in self.fields],
            'timestamp': datetime.now().isoformat()
        }
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        logging.info(f"Поля сохранены в {filename}")

    def load_fields(self, filename: str) -> bool:
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.fields = [FormField.from_dict(field_data) for field_data in data['fields']]
            logging.info(f"Загружено {len(self.fields)} полей из {filename}")
            return True
        except Exception as e:
            logging.error(f"Ошибка загрузки полей: {e}")
            return False


# ================== ОБРАБОТЧИК EXCEL ==================
class ExcelProcessor:
    @staticmethod
    def load_excel(filepath: str) -> Optional[pd.DataFrame]:
        try:
            df = pd.read_excel(filepath, header=None, dtype=str, engine='openpyxl')
            df = df.fillna('')
            df = df.applymap(lambda x: str(x).strip() if pd.notna(x) else '')
            logging.info(f"Загружен Excel файл: {filepath}, строк: {len(df)}")
            return df
        except Exception as e:
            logging.error(f"Ошибка загрузки Excel: {e}")
            return None

    @staticmethod
    def parse_date(date_str: str) -> Tuple[str, str, str]:
        if not date_str or pd.isna(date_str) or str(date_str).strip() == '':
            return '', '', ''
        date_str = str(date_str).strip()
        try:
            if ' ' in date_str:
                date_str = date_str.split()[0]
            formats = ['%Y-%m-%d', '%d.%m.%Y', '%d/%m/%Y', '%Y.%m.%d']
            dt = None
            for fmt in formats:
                try:
                    dt = datetime.strptime(date_str, fmt)
                    break
                except ValueError:
                    continue
            if dt is None:
                dt = date_parse(date_str, dayfirst=True)
            return f"{dt.day:02d}", f"{dt.month:02d}", str(dt.year)
        except Exception as e:
            logging.error(f"Ошибка парсинга даты '{date_str}': {e}")
            return date_str, date_str, date_str

    @staticmethod
    def extract_row_data(row: pd.Series) -> Dict[str, str]:
        data = {}
        row_list = row.tolist()
        data[FieldType.LAST_NAME] = row_list[1] if len(row_list) > 1 else ''
        data[FieldType.FIRST_NAME] = row_list[2] if len(row_list) > 2 else ''
        data[FieldType.MIDDLE_NAME] = row_list[3] if len(row_list) > 3 else ''
        if len(row_list) > 4:
            day, month, year = ExcelProcessor.parse_date(row_list[4])
            data[FieldType.BIRTH_DAY] = day
            data[FieldType.BIRTH_MONTH] = month
            data[FieldType.BIRTH_YEAR] = year
        else:
            data[FieldType.BIRTH_DAY] = ''
            data[FieldType.BIRTH_MONTH] = ''
            data[FieldType.BIRTH_YEAR] = ''
        return data


# ================== АВТОМАТИЗАТОР ==================
class Automator:
    def __init__(self, form_manager: FormManager):
        self.form_manager = form_manager
        self.is_running = False
        self.is_paused = False
        self.current_row = 0
        self.total_rows = 0
        self.df: Optional[pd.DataFrame] = None
        self.message_queue = queue.Queue()
        self.config = Config()
        self.setup_hotkeys()

    def setup_hotkeys(self):
        try:
            keyboard.add_hotkey('f1', self.toggle_pause)
            keyboard.add_hotkey('f2', self.stop)
        except:
            pass

    def toggle_pause(self):
        self.is_paused = not self.is_paused
        status = "приостановлена" if self.is_paused else "продолжена"
        self.message_queue.put(f"Автоматизация {status}")

    def stop(self):
        self.is_running = False
        self.message_queue.put("Автоматизация остановлена")

    def load_excel(self, filepath: str) -> bool:
        self.df = ExcelProcessor.load_excel(filepath)
        if self.df is not None:
            self.total_rows = len(self.df)
            return True
        return False

    def run(self, start_row: int = 0, speed_factor: float = 1.0) -> bool:
        if not self.form_manager.fields:
            self.message_queue.put("Ошибка: Сначала определите поля формы")
            return False
        if self.df is None:
            self.message_queue.put("Ошибка: Сначала загрузите Excel файл")
            return False
        if start_row >= self.total_rows or start_row < 0:
            self.message_queue.put("Ошибка: Некорректный номер стартовой строки")
            return False
        self.is_running = True
        self.is_paused = False
        self.current_row = start_row
        self.config.speed_factor = speed_factor
        thread = threading.Thread(target=self._run_automation, daemon=True)
        thread.start()
        return True

    def _run_automation(self):
        try:
            self.message_queue.put("Автоматизация начинается через 5 секунд...")
            time.sleep(5)
            for i in range(self.current_row, self.total_rows):
                if not self.is_running:
                    break
                while self.is_paused and self.is_running:
                    time.sleep(0.1)
                self.process_row(i)
                if i < self.total_rows - 1 and self.is_running:
                    time.sleep(1.0)
            if self.is_running:
                self.message_queue.put("✅ Автоматизация успешно завершена!")
            else:
                self.message_queue.put("⏹ Автоматизация остановлена")
        except Exception as e:
            self.message_queue.put(f"❌ Ошибка автоматизации: {str(e)}")
        finally:
            self.is_running = False

    def process_row(self, row_index: int):
        try:
            row = self.df.iloc[row_index]
            data = ExcelProcessor.extract_row_data(row)
            self.message_queue.put(
                f"📝 Обработка строки {row_index + 1}: {data[FieldType.LAST_NAME]} {data[FieldType.FIRST_NAME]}"
            )
            for field in self.form_manager.fields:
                if not self.is_running:
                    break
                value = data.get(field.field_type, '')
                if not value:
                    continue
                action = FormAction(field=field, value=value)
                success = action.execute(
                    self.config.speed_factor,
                    use_image=self.config.use_image_recognition,
                    verify=self.config.verify_input,
                    max_attempts=self.config.max_attempts,
                    ocr_lang=self.config.ocr_lang,
                    use_clipboard=self.config.use_clipboard
                )
                if not success:
                    self.message_queue.put(f"❌ Ошибка заполнения поля {field.name} в строке {row_index + 1}")
                    self.message_queue.put("Остановка автоматизации")
                    self.is_running = False
                    return
            self.message_queue.put(f"✅ Строка {row_index + 1} успешно обработана")
            time.sleep(0.5)
        except Exception as e:
            self.message_queue.put(f"❌ Критическая ошибка в строке {row_index + 1}: {str(e)}")
            self.is_running = False


# ================== ГРАФИЧЕСКИЙ ИНТЕРФЕЙС ==================
class SimpleGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("ВСУ СК РФ")
        self.root.geometry("900x700")
        self.config = Config.load()
        self.form_manager = FormManager()
        self.automator = Automator(self.form_manager)
        self.excel_path_var = tk.StringVar(value=self.config.excel_file)
        self.start_row_var = tk.IntVar(value=self.config.start_row + 1)
        self.speed_var = tk.DoubleVar(value=self.config.speed_factor)
        self.use_image_var = tk.BooleanVar(value=self.config.use_image_recognition)
        self.verify_input_var = tk.BooleanVar(value=self.config.verify_input)
        self.max_attempts_var = tk.IntVar(value=self.config.max_attempts)
        self.use_clipboard_var = tk.BooleanVar(value=self.config.use_clipboard)
        self.setup_ui()
        self.process_message_queue()
        self.setup_recording_hotkeys()

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        row = 0
        title_label = ttk.Label(main_frame, text="ВСУ СК РФ", font=("Arial", 16, "bold"))
        title_label.grid(row=row, column=0, columnspan=3, pady=(0, 10))
        row += 1

        excel_frame = ttk.LabelFrame(main_frame, text="Excel файл", padding="10")
        excel_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        excel_frame.columnconfigure(1, weight=1)
        ttk.Label(excel_frame, text="Путь к файлу:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        ttk.Entry(excel_frame, textvariable=self.excel_path_var, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E),
                                                                                padx=(0, 5))
        ttk.Button(excel_frame, text="Обзор", command=self.browse_excel).grid(row=0, column=2)
        ttk.Label(excel_frame, text="Начать с строки:").grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(10, 0))
        ttk.Spinbox(excel_frame, from_=1, to=100000, textvariable=self.start_row_var, width=10).grid(row=1, column=1,
                                                                                                     sticky=tk.W,
                                                                                                     pady=(10, 0))
        row += 1

        fields_frame = ttk.LabelFrame(main_frame, text="Поля формы", padding="10")
        fields_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        fields_buttons = ttk.Frame(fields_frame)
        fields_buttons.grid(row=0, column=0, columnspan=2, pady=(0, 10))
        self.record_btn = ttk.Button(fields_buttons, text="Начать запись полей", command=self.start_recording_fields,
                                     width=20)
        self.record_btn.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(fields_buttons, text="Сохранить поля", command=self.save_fields, width=15).pack(side=tk.LEFT,
                                                                                                   padx=(0, 5))
        ttk.Button(fields_buttons, text="Загрузить поля", command=self.load_fields, width=15).pack(side=tk.LEFT)

        options_frame = ttk.Frame(fields_frame)
        options_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        ttk.Checkbutton(options_frame, text="Распознавание по изображению", variable=self.use_image_var).pack(
            side=tk.LEFT, padx=(0, 10))
        ttk.Checkbutton(options_frame, text="Проверять ввод", variable=self.verify_input_var).pack(side=tk.LEFT,
                                                                                                   padx=(0, 10))
        ttk.Label(options_frame, text="Попыток:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Spinbox(options_frame, from_=1, to=10, textvariable=self.max_attempts_var, width=5).pack(side=tk.LEFT)
        ttk.Checkbutton(options_frame, text="Использовать буфер обмена", variable=self.use_clipboard_var).pack(
            side=tk.LEFT, padx=(10, 0))

        self.record_info = ttk.Label(fields_frame, text="Статус: Не записывается", foreground="gray")
        self.record_info.grid(row=2, column=0, columnspan=2, pady=(5, 0))
        row += 1

        auto_frame = ttk.LabelFrame(main_frame, text="Автоматизация", padding="10")
        auto_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        ttk.Label(auto_frame, text="Скорость:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        ttk.Scale(auto_frame, from_=0.5, to=3.0, variable=self.speed_var, length=200, orient=tk.HORIZONTAL).grid(row=0,
                                                                                                                 column=1,
                                                                                                                 sticky=tk.W,
                                                                                                                 padx=(
                                                                                                                     0,
                                                                                                                     20))
        self.speed_label = ttk.Label(auto_frame, text=f"{self.speed_var.get():.1f}x")
        self.speed_label.grid(row=0, column=2, sticky=tk.W)

        def update_speed_label(*args):
            self.speed_label.config(text=f"{self.speed_var.get():.1f}x")

        self.speed_var.trace_add("write", update_speed_label)

        auto_buttons = ttk.Frame(auto_frame)
        auto_buttons.grid(row=1, column=0, columnspan=3, pady=(10, 0))
        self.start_btn = ttk.Button(auto_buttons, text="Начать заполнение", command=self.start_automation, width=20)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.pause_btn = ttk.Button(auto_buttons, text="Пауза", command=self.toggle_automation_pause, width=15,
                                    state=tk.DISABLED)
        self.pause_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.stop_btn = ttk.Button(auto_buttons, text="Стоп", command=self.stop_automation, width=15, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT)
        row += 1

        log_frame = ttk.LabelFrame(main_frame, text="Логи", padding="10")
        log_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, font=("Consolas", 9))
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        log_buttons = ttk.Frame(log_frame)
        log_buttons.grid(row=1, column=0, sticky=tk.E, pady=(5, 0))
        ttk.Button(log_buttons, text="Очистить логи", command=self.clear_logs).pack(side=tk.RIGHT)
        row += 1

        self.status_var = tk.StringVar(value="Готов к работе")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_recording_hotkeys(self):
        self.recording_hotkeys = {
            '1': FieldType.LAST_NAME,
            '2': FieldType.FIRST_NAME,
            '3': FieldType.MIDDLE_NAME,
            '4': FieldType.BIRTH_DAY,
            '5': FieldType.BIRTH_MONTH,
            '6': FieldType.BIRTH_YEAR,
        }

    def browse_excel(self):
        filename = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.excel_path_var.set(filename)
            self.log_message(f"Выбран файл: {os.path.basename(filename)}")

    def start_recording_fields(self):
        if self.form_manager.is_recording:
            return
        use_image = self.use_image_var.get()
        self.form_manager.start_recording(use_image=use_image)
        self.record_btn.config(state=tk.DISABLED)
        self.record_info.config(text="Статус: Запись активна. Используйте клавиши 1-6 для записи полей",
                                foreground="red")
        self.log_message("🎬 Начата запись полей формы")
        self.log_message("Инструкция:")
        self.log_message(" 1. Переключитесь на окно с формой")
        self.log_message(" 2. Наведите курсор на САМУЮ ЛЕВУЮ ВЕРХНЮЮ ТОЧКУ поля 'Фамилия' и нажмите 1")
        self.log_message(" 3. Наведите на САМУЮ ЛЕВУЮ ВЕРХНЮЮ ТОЧКУ поля 'Имя' и нажмите 2")
        self.log_message(" 4. Повторите для всех полей (3-6)")
        self.log_message(" 8. Нажмите ESC для завершения записи")
        self.root.after(100, self.check_recording_keys)

    def check_recording_keys(self):
        if not self.form_manager.is_recording:
            return
        try:
            for key, field_type in self.recording_hotkeys.items():
                if keyboard.is_pressed(key):
                    x, y = pyautogui.position()
                    self.form_manager.record_field(field_type, (x, y))
                    self.log_message(f"📝 Записано поле: {field_type} на позиции ({x}, {y})")
                    time.sleep(0.5)
            if keyboard.is_pressed('esc'):
                self.form_manager.stop_recording()
                self.record_btn.config(state=tk.NORMAL)
                self.record_info.config(text="Статус: Запись завершена", foreground="green")
                self.log_message(f"✅ Запись завершена. Записано полей: {len(self.form_manager.fields)}")
            else:
                self.root.after(50, self.check_recording_keys)
        except Exception as e:
            self.log_message(f"Ошибка записи: {e}")
            self.form_manager.stop_recording()
            self.record_btn.config(state=tk.NORMAL)
            self.record_info.config(text="Статус: Ошибка записи", foreground="red")

    def save_fields(self):
        if not self.form_manager.fields:
            messagebox.showwarning("Предупреждение", "Нет записанных полей для сохранения")
            return
        filename = filedialog.asksaveasfilename(
            title="Сохранить поля формы",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            self.form_manager.save_fields(filename)
            self.log_message(f"💾 Поля сохранены в {filename}")

    def load_fields(self):
        filename = filedialog.askopenfilename(
            title="Загрузить поля формы",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            if self.form_manager.load_fields(filename):
                self.log_message(f"📂 Загружено {len(self.form_manager.fields)} полей")
            else:
                messagebox.showerror("Ошибка", "Не удалось загрузить поля")

    def start_automation(self):
        self.config.excel_file = self.excel_path_var.get()
        self.config.start_row = self.start_row_var.get() - 1
        self.config.speed_factor = self.speed_var.get()
        self.config.use_image_recognition = self.use_image_var.get()
        self.config.verify_input = self.verify_input_var.get()
        self.config.max_attempts = self.max_attempts_var.get()
        self.config.use_clipboard = self.use_clipboard_var.get()

        if not HAS_PYPERCLIP and self.config.use_clipboard:
            response = messagebox.askyesno("pyperclip не установлен",
                                           "Модуль pyperclip не установлен. Без него использование буфера обмена невозможно.\n"
                                           "Установите его командой: pip install pyperclip\n\n"
                                           "Продолжить без использования буфера обмена?")
            if response:
                self.config.use_clipboard = False
            else:
                return

        if not self.config.excel_file or not os.path.exists(self.config.excel_file):
            messagebox.showerror("Ошибка", "Выберите существующий Excel файл")
            return
        if not self.form_manager.fields:
            messagebox.showerror("Ошибка", "Сначала определите или загрузите поля формы")
            return
        if not self.automator.load_excel(self.config.excel_file):
            messagebox.showerror("Ошибка", "Не удалось загрузить Excel файл")
            return
        if messagebox.askyesno("Подтверждение",
                               "Запустить автоматизацию?\n\nУбедитесь, что:\n" "1. Форма открыта и видна\n" "2. Курсор мыши можно переместить в левый верхний угол для остановки"):
            self.config.save()
            if self.automator.run(self.config.start_row, self.config.speed_factor):
                self.start_btn.config(state=tk.DISABLED)
                self.pause_btn.config(state=tk.NORMAL)
                self.stop_btn.config(state=tk.NORMAL)
                self.status_var.set("Автоматизация запущена")
                self.log_message("▶ Запущена автоматизация")
            else:
                messagebox.showerror("Ошибка", "Не удалось запустить автоматизацию")

    def toggle_automation_pause(self):
        self.automator.toggle_pause()
        if self.automator.is_paused:
            self.pause_btn.config(text="Продолжить")
        else:
            self.pause_btn.config(text="Пауза")

    def stop_automation(self):
        self.automator.stop()
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.DISABLED)
        self.pause_btn.config(text="Пауза")

    def log_message(self, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)

    def clear_logs(self):
        self.log_text.delete(1.0, tk.END)

    def process_message_queue(self):
        try:
            while True:
                message = self.automator.message_queue.get_nowait()
                self.log_message(message)
                if "остановлена" in message.lower():
                    self.status_var.set("Автоматизация остановлена")
                    self.start_btn.config(state=tk.NORMAL)
                    self.pause_btn.config(state=tk.DISABLED)
                    self.stop_btn.config(state=tk.DISABLED)
                elif "завершена" in message.lower():
                    self.status_var.set("Автоматизация завершена")
                    self.start_btn.config(state=tk.NORMAL)
                    self.pause_btn.config(state=tk.DISABLED)
                    self.stop_btn.config(state=tk.DISABLED)
                elif "ошибка" in message.lower():
                    self.status_var.set("Ошибка")
        except queue.Empty:
            pass
        self.root.after(100, self.process_message_queue)

    def on_closing(self):
        self.config.save()
        self.automator.stop()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


# ================== ТОЧКА ВХОДА ==================
def main():
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('auto_form_filler.log', encoding='utf-8'),
            logging.StreamHandler()
        ]
    )

    # Проверка необходимых библиотек
    try:
        import pandas as pd
        import pyautogui
        import keyboard
        from dateutil.parser import parse
        from PIL import Image, ImageGrab
        import pytesseract
        # pyperclip проверяется выше
    except ImportError as e:
        print(f"Ошибка: Не установлена необходимая библиотека: {e}")
        print("Установите библиотеки командой:")
        print("pip install pandas openpyxl pyautogui keyboard python-dateutil pillow pytesseract pyperclip")
        print("Также установите Tesseract OCR:")
        print("  - Ubuntu/Debian: sudo apt install tesseract-ocr tesseract-ocr-rus")
        print("  - Windows: скачайте установщик с https://github.com/UB-Mannheim/tesseract/wiki")
        print("  - MacOS: brew install tesseract tesseract-lang")
        input("Нажмите Enter для выхода...")
        return

    # Проверка наличия Tesseract
    try:
        pytesseract.get_tesseract_version()
    except Exception as e:
        print(f"Предупреждение: Tesseract OCR не найден: {e}")
        print("Для работы OCR необходимо установить Tesseract:")
        print("  - Ubuntu/Debian: sudo apt install tesseract-ocr tesseract-ocr-rus")
        print("  - Windows: скачайте установщик с https://github.com/UB-Mannheim/tesseract/wiki")
        print("  - MacOS: brew install tesseract tesseract-lang")
        print("Продолжаем без OCR...")

    app = SimpleGUI()
    try:
        app.run()
    except Exception as e:
        logging.error(f"Критическая ошибка: {e}")
        messagebox.showerror("Ошибка", f"Критическая ошибка: {e}")


if __name__ == "__main__":
    main()