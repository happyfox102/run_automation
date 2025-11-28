import sys
import time
import pickle
import os
from pathlib import Path
import pandas as pd
import pyautogui
import pyperclip
from pynput import mouse
from datetime import datetime, timedelta
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QPushButton,
    QVBoxLayout, QFileDialog, QTextEdit, QLabel, QComboBox, QHBoxLayout, QSpinBox
)
import threading

# ================= НАСТРОЙКИ =================
pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.05
ACTIONS_FILE = "actions.pkl"
actions = []
recording = False
m_listener = None
running = False
EXCEL_FILE = None
df = None
SPEED_FACTOR = 1.0
window = None
START_ROW = 0  # пользователь может выбрать с какой строки начинать (0-indexed)

# =================== УТИЛИТЫ ===================

def safe_sleep(sec):
    """Небольшая обёртка для time.sleep, учитывающая SPEED_FACTOR."""
    time.sleep(max(0, sec * SPEED_FACTOR))


def clear_text_field_improved():
    """Более надёжное очищение поля:
    - гарантируем, что поле выделено (предполагается, что перед этим был клик по полю);
    - используем Ctrl+A + Backspace + Delete;
    - небольшой таймаут для стабильности.
    """
    try:
        # Выделяем всё
        pyautogui.hotkey('ctrl', 'a')
        safe_sleep(0.03)
        # Удаляем содержимое несколькими вариантами (на случай если одно не сработает)
        pyautogui.press('backspace')
        safe_sleep(0.02)
        pyautogui.press('delete')
        safe_sleep(0.02)
    except Exception:
        # Последняя мера — вставить пустую строку из буфера, это надежно заменит текст
        try:
            pyperclip.copy('')
            safe_sleep(0.01)
            pyautogui.hotkey('ctrl', 'v')
            safe_sleep(0.02)
        except Exception:
            pass


def paste_text_improved(text: str):
    """Надёжная вставка текста.
    Подходы:
    1) Копируем в буфер и вставляем Ctrl+V — быстро и корректно для любых символов.
    2) Если вставка через буфер почему-то не сработала, делаем ввод методом печати (typewrite).

    Также делаем небольшую проверку: если текст короткий (<60 знаков) — вводим медленно,
    чтобы избежать проблем с полями, которые реагируют по-особому (например автодополнение).
    """
    try:
        # Сначала попробуем через буфер обмена — это обычно самый быстрый и надёжный способ
        pyperclip.copy(str(text))
        safe_sleep(0.03)
        pyautogui.hotkey('ctrl', 'v')
        safe_sleep(0.04)
    except Exception:
        # Фоллбек — симулируем печать
        try:
            # Если текст очень длинный, печатаем быстрее, иначе — медленно
            interval = 0.01 if len(str(text)) > 60 else 0.03
            pyautogui.typewrite(str(text), interval=interval)
            safe_sleep(0.02)
        except Exception:
            # Если и это не проходит — ничего не делаем
            pass


# =================== UI ===================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Автозаполнение форм — улучшенная версия")
        self.setGeometry(200, 200, 600, 420)

        layout = QVBoxLayout()

        # Excel
        self.excel_label = QLabel("Excel: не загружен")
        layout.addWidget(self.excel_label)

        btn_row = QHBoxLayout()
        self.load_button = QPushButton("📂 Загрузить Excel")
        self.load_button.clicked.connect(self.load_excel)
        btn_row.addWidget(self.load_button)

        btn_row.addStretch()
        start_row_label = QLabel("Стартовая строка (1 = первая):")
        btn_row.addWidget(start_row_label)
        self.start_row_spin = QSpinBox()
        self.start_row_spin.setMinimum(1)
        self.start_row_spin.setMaximum(1000000)
        self.start_row_spin.setValue(1)
        self.start_row_spin.valueChanged.connect(self.update_start_row)
        btn_row.addWidget(self.start_row_spin)

        layout.addLayout(btn_row)

        # Запись
        self.record_button = QPushButton("🔴 Начать запись")
        self.record_button.clicked.connect(start_recording)
        layout.addWidget(self.record_button)

        self.stop_record_button = QPushButton("🟥 Остановить запись")
        self.stop_record_button.clicked.connect(stop_recording)
        layout.addWidget(self.stop_record_button)

        # Автоматизация
        self.start_button = QPushButton("▶ Запустить авто-заполнение")
        self.start_button.clicked.connect(self.start_automation_thread)
        layout.addWidget(self.start_button)

        self.stop_button = QPushButton("⛔ Остановить")
        self.stop_button.clicked.connect(stop_automation)
        layout.addWidget(self.stop_button)

        # Скорость
        speed_label = QLabel("⚡ Скорость работы:")
        layout.addWidget(speed_label)
        self.speed_box = QComboBox()
        self.speed_box.addItems([
            "Очень быстро (0.5)",
            "Быстро (1.0)",
            "Нормально (1.5)",
            "Медленно (2.0)",
            "Очень медленно (3.0)"
        ])
        self.speed_box.setCurrentIndex(1)
        self.speed_box.currentIndexChanged.connect(self.update_speed)
        layout.addWidget(self.speed_box)

        # Логи
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        layout.addWidget(self.log_box)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def log(self, text):
        self.log_box.append(text)
        QApplication.processEvents()

    def load_excel(self):
        global EXCEL_FILE, df
        file_path, _ = QFileDialog.getOpenFileName(self, "Выбор Excel файла", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            try:
                # Читаем весь лист как строки, чтобы не терять форматы и избежать преобразований
                df = pd.read_excel(file_path, header=None, dtype=str)
                df = df.fillna('')
                EXCEL_FILE = file_path
                self.excel_label.setText(f"✅ Загружен: {os.path.basename(file_path)} — строки: {len(df)}")
                self.log(f"📄 Excel загружен: {file_path}")
            except Exception as e:
                self.log(f"❌ Ошибка загрузки: {str(e)}")

    def update_speed(self):
        global SPEED_FACTOR
        speeds = {0: 0.5, 1: 1.0, 2: 1.5, 3: 2.0, 4: 3.0}
        SPEED_FACTOR = speeds.get(self.speed_box.currentIndex(), 1.0)
        self.log(f"⚡ Скорость установлена: {SPEED_FACTOR}")

    def start_automation_thread(self):
        global running
        if running:
            return
        thread = threading.Thread(target=run_automation, daemon=True)
        thread.start()

    def update_start_row(self, val):
        global START_ROW
        START_ROW = max(0, val - 1)


# ================== Запись кликов ==================

def on_click(x, y, button, pressed):
    global actions
    if recording and pressed:
        # Сохраняем тип кнопки и координаты и относительное время
        actions.append(('click', time.time(), x, y, str(button)))


def start_recording():
    global recording, actions, m_listener, window
    actions = []
    recording = True
    if window:
        window.log("🔴 Началась запись кликов (кликните по полям ввода по порядку)")
    m_listener = mouse.Listener(on_click=on_click)
    m_listener.start()


def stop_recording():
    global recording, m_listener, window
    recording = False
    if m_listener:
        m_listener.stop()
    try:
        with open(ACTIONS_FILE, 'wb') as f:
            pickle.dump(actions, f)
        if window:
            window.log(f"✅ Запись остановлена. Сохранено действий: {len(actions)}")
    except Exception as e:
        if window:
            window.log(f"❌ Ошибка при сохранении действий: {e}")


# ================== Автоматизация ==================

def stop_automation():
    global running, window
    running = False
    if window:
        window.log("🛑 Автоматизация остановлена вручную")


def run_automation():
    global running, actions, df, window, START_ROW
    if window is None:
        return
    if df is None:
        window.log("❌ Сначала загрузите Excel файл")
        return
    if not os.path.exists(ACTIONS_FILE):
        window.log("❌ Сначала запишите клики")
        return

    # Загружаем записанные клики
    try:
        with open(ACTIONS_FILE, 'rb') as f:
            actions = pickle.load(f)
    except Exception as e:
        window.log(f"❌ Не удалось прочитать {ACTIONS_FILE}: {e}")
        return

    window.log("⏳ 5 секунд на переход в приложение/браузер и размещение окна")
    time.sleep(5)
    running = True

    # Пробегаем по строкам Excel, начиная с START_ROW
    for idx in range(START_ROW, len(df)):
        if not running:
            break
        row = df.iloc[idx]

        # Формируем список значений для вставки (подстраиваемся под длину строки)
        try:
            # Преобразуем значения в строки и убираем лишние пробелы
            current_values = [str(v).strip() if v is not None else '' for v in row.tolist()]
            # Если в строке есть столбец с датой в 4-й позиции (именно как в вашем коде) — приводим к day/month/year
            if len(current_values) > 3 and current_values[3] != '':
                raw = current_values[3]
                # Пытаемся распознать как число-сериал Excel
                try:
                    serial = float(raw)
                    date = datetime(1899, 12, 30) + timedelta(days=serial)
                    current_values[3:4] = [f"{date.day:02d}", f"{date.month:02d}", str(date.year)]
                except Exception:
                    # Попробуем разобрать yyyy-mm-dd или dd.mm.yyyy
                    if '-' in raw:
                        try:
                            parts = raw.split()[0].split('-')
                            year, month, day = map(int, parts[:3])
                            current_values[3:4] = [f"{day:02d}", f"{month:02d}", str(year)]
                        except Exception:
                            pass
                    elif '.' in raw:
                        try:
                            parts = raw.split()[0].split('.')
                            day, month, year = map(int, parts[:3])
                            current_values[3:4] = [f"{day:02d}", f"{month:02d}", str(year)]
                        except Exception:
                            pass

            num_fields = len(current_values)
        except Exception as e:
            window.log(f"❌ Ошибка при обработке строки {idx+1}: {e}")
            continue

        # Пошагово воспроизводим записанные клики
        if actions:
            base_time = actions[0][1]
            start_time = time.time()
            field_index = 0

            for act in actions:
                if not running:
                    break
                if act[0] != 'click':
                    continue
                # Воспроизводим относительную задержку
                delay = max(0.0, (act[1] - base_time) * SPEED_FACTOR - (time.time() - start_time))
                if delay > 0:
                    time.sleep(delay)

                # Клик по координатам
                try:
                    x, y = int(act[2]), int(act[3])
                    pyautogui.click(x, y)
                    safe_sleep(0.06)

                    # Очищаем поле и вставляем значение
                    clear_text_field_improved()
                    # Если закончились значения, вставляем пустую строку
                    value = current_values[field_index] if field_index < num_fields else ''
                    field_index += 1

                    paste_text_improved(value)
                    safe_sleep(0.08)
                except Exception as e:
                    window.log(f"⚠️ Ошибка при клике/вставке: {e}")
                    continue

        window.log(f"✅ Строка {idx+1} вставлена")
        # Небольшая пауза между строками — чтобы веб-страница/приложение успело обработать ввод
        safe_sleep(1.0)

    window.log("🏁 Автоматизация завершена (один проход по Excel)")
    running = False


# ================== ЗАПУСК ==================
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
