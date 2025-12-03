import sys
import time
import pickle
import os
from pathlib import Path
import pandas as pd
import pyautogui
import pyperclip
from PyQt5 import Qt
from pynput import mouse
from datetime import datetime, timedelta
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QPushButton, QVBoxLayout,
    QFileDialog, QTextEdit, QLabel, QComboBox, QHBoxLayout, QSpinBox,
    QMessageBox
)
import threading

# ================= НАСТРОЙКИ =================
pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.1

ACTIONS_FILE = "actions.pkl"
actions = []
recording = False
m_listener = None
running = False
EXCEL_FILE = None
df = None
SPEED_FACTOR = 1.0
window = None
START_ROW = 0
PAUSE_BETWEEN_ROWS = 1.0


# ================= УТИЛИТЫ =================
def safe_sleep(sec):
    time.sleep(max(0, sec * SPEED_FACTOR))


def clear_text_field():
    """Очистка текстового поля"""
    try:
        pyautogui.hotkey('ctrl', 'a')
        safe_sleep(0.05)
        pyautogui.press('delete')
        safe_sleep(0.05)
    except:
        try:
            pyautogui.click(clicks=3)
            safe_sleep(0.05)
            pyautogui.press('backspace')
            safe_sleep(0.05)
        except:
            pass


def paste_text(text):
    """Вставка текста"""
    try:
        text = str(text).strip()
        pyperclip.copy(text)
        safe_sleep(0.05)
        pyautogui.hotkey('ctrl', 'v')
        safe_sleep(0.1)
        return True
    except:
        try:
            pyautogui.write(text, interval=0.01)
            safe_sleep(0.1)
            return True
        except:
            return False


def process_excel_date(date_value):
    """Обработка даты из Excel"""
    if pd.isna(date_value) or str(date_value).strip() == '':
        return ['', '', '']

    try:
        # Если это число (Excel serial date)
        try:
            excel_date = float(date_value)
            base_date = datetime(1899, 12, 30)
            date_obj = base_date + timedelta(days=excel_date)
            return [
                f"{date_obj.day:02d}",
                f"{date_obj.month:02d}",
                str(date_obj.year)
            ]
        except:
            pass

        # Если это строка
        date_str = str(date_value).strip()

        # Пробуем разные форматы
        formats = ['%Y-%m-%d', '%d.%m.%Y', '%d/%m/%Y', '%Y.%m.%d', '%m/%d/%Y']

        for fmt in formats:
            try:
                date_obj = datetime.strptime(date_str.split()[0], fmt)
                return [
                    f"{date_obj.day:02d}",
                    f"{date_obj.month:02d}",
                    str(date_obj.year)
                ]
            except:
                continue

        # Если не распарсилось, возвращаем как есть
        return [date_str, '', '']

    except Exception as e:
        return [str(date_value), '', '']


# ================= ГЛАВНОЕ ОКНО =================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Автозаполнение форм")
        self.setGeometry(200, 200, 600, 500)

        layout = QVBoxLayout()

        # 1. Загрузка Excel
        excel_group = QWidget()
        excel_layout = QVBoxLayout(excel_group)

        self.excel_label = QLabel("Excel файл не загружен")
        excel_layout.addWidget(self.excel_label)

        excel_btn_layout = QHBoxLayout()
        self.load_button = QPushButton("📂 Загрузить Excel")
        self.load_button.clicked.connect(self.load_excel)
        excel_btn_layout.addWidget(self.load_button)

        excel_btn_layout.addWidget(QLabel("Старт с строки:"))
        self.start_row_spin = QSpinBox()
        self.start_row_spin.setMinimum(1)
        self.start_row_spin.setMaximum(10000)
        self.start_row_spin.setValue(1)
        self.start_row_spin.valueChanged.connect(self.update_start_row)
        excel_btn_layout.addWidget(self.start_row_spin)

        excel_layout.addLayout(excel_btn_layout)
        layout.addWidget(excel_group)

        # 2. Запись действий
        record_group = QWidget()
        record_layout = QVBoxLayout(record_group)

        self.record_info = QLabel(f"Записано действий: {len(actions)}")
        record_layout.addWidget(self.record_info)

        record_btn_layout = QHBoxLayout()
        self.record_button = QPushButton("🔴 Начать запись")
        self.record_button.clicked.connect(self.start_recording)
        record_btn_layout.addWidget(self.record_button)

        self.stop_record_button = QPushButton("■ Остановить запись")
        self.stop_record_button.clicked.connect(self.stop_recording)
        self.stop_record_button.setEnabled(False)
        record_btn_layout.addWidget(self.stop_record_button)

        self.clear_actions_button = QPushButton("🗑️ Очистить")
        self.clear_actions_button.clicked.connect(self.clear_actions)
        record_btn_layout.addWidget(self.clear_actions_button)

        record_layout.addLayout(record_btn_layout)
        layout.addWidget(record_group)

        # 3. Автоматизация
        auto_group = QWidget()
        auto_layout = QVBoxLayout(auto_group)

        self.status_label = QLabel("Статус: Готов")
        auto_layout.addWidget(self.status_label)

        auto_btn_layout = QHBoxLayout()
        self.start_button = QPushButton("▶ Запустить заполнение")
        self.start_button.clicked.connect(self.start_automation)
        auto_btn_layout.addWidget(self.start_button)

        self.stop_button = QPushButton("⏹ Остановить")
        self.stop_button.clicked.connect(self.stop_automation)
        self.stop_button.setEnabled(False)
        auto_btn_layout.addWidget(self.stop_button)

        auto_layout.addLayout(auto_btn_layout)
        layout.addWidget(auto_group)

        # 4. Настройки
        settings_group = QWidget()
        settings_layout = QVBoxLayout(settings_group)

        # Скорость
        speed_layout = QHBoxLayout()
        speed_layout.addWidget(QLabel("Скорость:"))
        self.speed_combo = QComboBox()
        self.speed_combo.addItems([
            "Очень быстро (0.5)",
            "Быстро (0.8)",
            "Нормально (1.0)",
            "Медленно (1.5)",
            "Очень медленно (2.0)"
        ])
        self.speed_combo.setCurrentIndex(2)
        self.speed_combo.currentIndexChanged.connect(self.update_speed)
        speed_layout.addWidget(self.speed_combo)
        settings_layout.addLayout(speed_layout)

        # Паузы
        delay_layout = QHBoxLayout()
        delay_layout.addWidget(QLabel("Пауза между строками:"))
        self.pause_spin = QSpinBox()
        self.pause_spin.setRange(1, 10)
        self.pause_spin.setValue(1)
        self.pause_spin.setSuffix(" сек")
        self.pause_spin.valueChanged.connect(self.update_pause)
        delay_layout.addWidget(self.pause_spin)
        settings_layout.addLayout(delay_layout)

        layout.addWidget(settings_group)

        # 5. Лог
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        layout.addWidget(self.log_box)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # Загружаем сохраненные действия
        self.load_actions()

    def log(self, text):
        self.log_box.append(text)
        QApplication.processEvents()

    def load_excel(self):
        global EXCEL_FILE, df
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите Excel файл", "",
            "Excel Files (*.xlsx *.xls);;All Files (*.*)"
        )
        if file_path:
            try:
                df = pd.read_excel(file_path, header=None, dtype=str)
                df = df.fillna('')
                EXCEL_FILE = file_path
                self.excel_label.setText(f"✅ Загружен: {os.path.basename(file_path)} ({len(df)} строк)")
                self.log(f"📄 Загружен Excel файл: {len(df)} строк")
            except Exception as e:
                self.log(f"❌ Ошибка загрузки: {str(e)}")
                QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить файл:\n{str(e)}")

    def update_speed(self):
        global SPEED_FACTOR
        speeds = {0: 0.5, 1: 0.8, 2: 1.0, 3: 1.5, 4: 2.0}
        SPEED_FACTOR = speeds.get(self.speed_combo.currentIndex(), 1.0)
        self.log(f"⚡ Установлена скорость: {SPEED_FACTOR}")

    def update_start_row(self, val):
        global START_ROW
        START_ROW = max(0, val - 1)

    def update_pause(self, val):
        global PAUSE_BETWEEN_ROWS
        PAUSE_BETWEEN_ROWS = float(val)

    def start_recording(self):
        thread = threading.Thread(target=start_recording, daemon=True)
        thread.start()

    def stop_recording(self):
        thread = threading.Thread(target=stop_recording, daemon=True)
        thread.start()

    def clear_actions(self):
        global actions
        actions = []
        if os.path.exists(ACTIONS_FILE):
            os.remove(ACTIONS_FILE)
        self.record_info.setText(f"Записано действий: 0")
        self.log("🗑️ Все действия очищены")

    def load_actions(self):
        global actions
        if os.path.exists(ACTIONS_FILE):
            try:
                with open(ACTIONS_FILE, 'rb') as f:
                    actions = pickle.load(f)
                self.record_info.setText(f"Записано действий: {len(actions)}")
                self.log(f"📝 Загружено {len(actions)} сохраненных действий")
            except Exception as e:
                self.log(f"⚠️ Ошибка загрузки действий: {e}")

    def start_automation(self):
        global running
        if running:
            return
        thread = threading.Thread(target=run_automation, daemon=True)
        thread.start()

    def stop_automation(self):
        global running
        running = False
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.status_label.setText("Статус: Остановлено")
        self.log("🛑 Автоматизация остановлена")


# ================= ЗАПИСЬ ДЕЙСТВИЙ =================
def on_click(x, y, button, pressed):
    global actions
    if recording and pressed and button == mouse.Button.left:
        actions.append(('click', time.time(), x, y))
        if window:
            window.record_info.setText(f"Записано действий: {len(actions)}")


def start_recording():
    global recording, m_listener
    actions.clear()
    recording = True

    if window:
        window.record_button.setEnabled(False)
        window.stop_record_button.setEnabled(True)
        window.log("🔴 Запись начата! Кликайте по полям ЛЕВОЙ кнопкой мыши.")
        window.log("Нажмите 'Остановить запись' когда закончите.")

    m_listener = mouse.Listener(on_click=on_click)
    m_listener.start()


def stop_recording():
    global recording, m_listener
    recording = False

    if m_listener:
        m_listener.stop()
        m_listener = None

    try:
        with open(ACTIONS_FILE, 'wb') as f:
            pickle.dump(actions, f)
        if window:
            window.record_button.setEnabled(True)
            window.stop_record_button.setEnabled(False)
            window.log(f"✅ Запись остановлена. Сохранено {len(actions)} действий")
    except Exception as e:
        if window:
            window.log(f"❌ Ошибка сохранения действий: {e}")


# ================= АВТОМАТИЗАЦИЯ =================
def run_automation():
    global running, df, window

    if window is None:
        return

    if df is None or df.empty:
        window.log("❌ Ошибка: Сначала загрузите Excel файл!")
        QMessageBox.warning(window, "Ошибка", "Сначала загрузите Excel файл!")
        return

    if not actions:
        window.log("❌ Ошибка: Сначала запишите действия!")
        QMessageBox.warning(window, "Ошибка", "Сначала запишите действия кликами по полям!")
        return

    window.log("⏱️ Подготовка... У вас 5 секунд чтобы открыть форму!")
    window.start_button.setEnabled(False)
    window.stop_button.setEnabled(True)
    window.status_label.setText("Статус: Выполняется...")

    time.sleep(5)

    running = True

    try:
        for row_idx in range(START_ROW, len(df)):
            if not running:
                break

            row = df.iloc[row_idx]
            window.status_label.setText(f"Статус: Строка {row_idx + 1}/{len(df)}")

            # Получаем значения строки
            current_values = [str(v).strip() if pd.notna(v) else '' for v in row.tolist()]

            # Обрабатываем дату (4-й столбец, индекс 3)
            if len(current_values) > 3 and current_values[3]:
                date_parts = process_excel_date(current_values[3])
                # Заменяем дату на три отдельных поля
                current_values = current_values[:3] + list(date_parts) + current_values[4:]
                if any(date_parts):
                    window.log(f"📅 Строка {row_idx + 1}: дата разделена")

            window.log(f"📝 Обработка строки {row_idx + 1}")

            # Выполняем записанные действия
            prev_time = actions[0][1] if actions else time.time()
            field_index = 0

            for action in actions:
                if not running:
                    break

                if action[0] != 'click':
                    continue

                # Рассчет задержки
                recorded_delay = action[1] - prev_time
                adjusted_delay = max(0, recorded_delay * SPEED_FACTOR)
                elapsed = time.time() - prev_time
                sleep_time = max(0, adjusted_delay - elapsed)
                if sleep_time > 0:
                    time.sleep(sleep_time)

                prev_time = action[1]

                # Клик и заполнение поля
                try:
                    x, y = action[2], action[3]

                    # Кликаем в поле
                    pyautogui.moveTo(x, y, duration=0.1)
                    pyautogui.click(x, y)
                    safe_sleep(0.1)

                    # F2 для редактирования (если это Excel/таблица)
                    pyautogui.press('f2')
                    safe_sleep(0.1)

                    # Очищаем поле
                    clear_text_field()
                    safe_sleep(0.1)

                    # Получаем значение для вставки
                    if field_index < len(current_values):
                        value = current_values[field_index]
                    else:
                        value = ''

                    # ВСТАВКА ЗНАЧЕНИЯ (исправлено!)
                    if value:
                        success = paste_text(value)
                        if success:
                            window.log(f"  ✓ Поле {field_index + 1}: '{value}'")
                        else:
                            window.log(f"  ✗ Поле {field_index + 1}: ошибка вставки")
                    else:
                        window.log(f"  ∅ Поле {field_index + 1}: пусто")

                    field_index += 1

                except Exception as e:
                    window.log(f"  ⚠️ Ошибка в поле {field_index + 1}: {str(e)}")
                    field_index += 1
                    continue

            # Пауза между строками
            if running and row_idx < len(df) - 1:
                window.log(f"⏸ Пауза {PAUSE_BETWEEN_ROWS} сек...")
                safe_sleep(PAUSE_BETWEEN_ROWS)

        if running:
            window.log("✅ Заполнение завершено!")
            QMessageBox.information(window, "Успех", "Заполнение формы завершено!")

    except Exception as e:
        window.log(f"❌ Критическая ошибка: {str(e)}")
        QMessageBox.critical(window, "Ошибка", f"Произошла ошибка:\n{str(e)}")

    finally:
        running = False
        window.start_button.setEnabled(True)
        window.stop_button.setEnabled(False)
        window.status_label.setText("Статус: Готов")


# ================= ЗАПУСК =================
if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    # Настройки для Windows
    if hasattr(QApplication, 'setAttribute'):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())