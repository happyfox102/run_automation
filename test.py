import pyautogui
import time
import pickle
import keyboard
import pandas as pd
import os
from pynput import mouse
import pyperclip
import webbrowser  # Добавлен для открытия браузера

pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.02

# URL формы — замените на актуальный URL вашей формы
FORM_URL = "https://your.form.url"  # Замените на URL формы

# ========= ПОИСК EXCEL =========
def find_excel_file():
    for f in os.listdir():
        if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~$'):
            return f
    return None

EXCEL_FILE = find_excel_file()
if not EXCEL_FILE:
    print("❌ Не найден Excel файл")
    input()
    exit()

print(f"✅ Найден файл: {EXCEL_FILE}")

# ========= ЧТЕНИЕ =========
df = pd.read_excel(EXCEL_FILE, header=None)

# ========= ГЛОБАЛЬНЫЕ =========
ACTIONS_FILE = "actions.pkl"
actions = []
recording = False
m_listener = None

# ========= НАДЁЖНОЕ ОЧИЩЕНИЕ =========
def clear_text_field():
    time.sleep(0.05)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.05)
    pyautogui.press('delete')
    time.sleep(0.05)

# ========= СНЯТИЕ ЧЕКБОКСОВ =========
def clear_checkboxes():
    try:
        while True:
            box = pyautogui.locateCenterOnScreen("checkbox_checked.png", confidence=0.8)
            if not box:
                break
            pyautogui.click(box.x, box.y)
            time.sleep(0.2)
    except:
        pass

# ========= ЗАПИСЬ =========
def on_click(x, y, button, pressed):
    global actions
    if recording and pressed:
        actions.append(('click', time.time(), x, y))

def start_recording():
    global recording, actions, m_listener
    actions = []
    recording = True
    print("🔴 Запись началась. Записываются только клики. Кликайте по полям формы в порядке: фамилия, имя, отчество, номер, затем по другим элементам (кнопки submit и т.д.) если нужно.")

    m_listener = mouse.Listener(on_click=on_click)
    m_listener.start()

def stop_recording():
    global recording, m_listener
    recording = False

    if m_listener: m_listener.stop()

    with open(ACTIONS_FILE, 'wb') as f:
        pickle.dump(actions, f)

    print(f"✅ Сохранено действий: {len(actions)}")

# ========= АВТО =========
def run_automation():
    global actions

    if not actions:
        try:
            with open(ACTIONS_FILE, 'rb') as f:
                actions = pickle.load(f)
        except:
            print("❌ Нет шаблона")
            return

    print("\n⏳ Открываем браузер и даем 5 секунд на загрузку")
    webbrowser.open(FORM_URL)
    time.sleep(5)

    # Инициализация маппинга полей на основе плейсхолдеров (только для первой итерации)
    field_mapping = {}  # ключ: индекс акта, значение: индекс значения (0: фамилия, 1: имя, etc.)
    previous_values = [None] * 4  # Для хранения предыдущих значений

    iteration = 0

    while True:  # Бесконечный цикл
        for i, row in df.iterrows():
            try:
                current_values = [
                    str(row[0]).strip(),
                    str(row[1]).strip(),
                    str(row[2]).strip(),
                    str(row[3]).strip()
                ]
            except:
                continue

            print(f"\n▶ Итерация {iteration + 1}, Строка {i+1}")

            clear_checkboxes()

            text_index = 0
            start_time = time.time()
            base_time = actions[0][1] if actions else 0

            for j, act in enumerate(actions):
                delay = act[1] - base_time
                passed = time.time() - start_time
                if delay > passed:
                    time.sleep(delay - passed)

                if act[0] == 'click':
                    pyautogui.click(act[2], act[3])
                    time.sleep(0.1)

                    old_clip = pyperclip.paste()
                    pyperclip.copy("%%KNOWN%%")
                    time.sleep(0.05)
                    pyautogui.hotkey('ctrl', 'a')
                    time.sleep(0.05)
                    pyautogui.hotkey('ctrl', 'c')
                    time.sleep(0.05)
                    current_text = pyperclip.paste().strip()

                    if current_text != "%%KNOWN%%":
                        # Это текстовое поле
                        if iteration == 0:
                            # Первая итерация: определяем маппинг по плейсхолдерам
                            placeholders = ["{ФАМИЛИЯ}", "{ИМЯ}", "{ОТЧЕСТВО}", "{НОМЕР}"]
                            lower_placeholders = ["{фамилия}", "{имя}", "{отчество}", "{номер}"]
                            for ph_idx, ph in enumerate(placeholders + lower_placeholders):
                                if current_text.lower() == ph.lower():
                                    field_mapping[j] = ph_idx % 4  # 0-3
                                    break
                            else:
                                # Если не плейсхолдер, пропустить
                                continue

                        # Получаем индекс поля из маппинга
                        if j in field_mapping:
                            field_idx = field_mapping[j]
                            # Если есть предыдущее значение, проверяем, совпадает ли current_text с previous_values[field_idx]
                            if previous_values[field_idx] and current_text == previous_values[field_idx]:
                                # Заменяем на новое
                                clear_text_field()
                                to_paste = current_values[field_idx]
                                pyperclip.copy(to_paste)
                                time.sleep(0.05)
                                pyautogui.hotkey('ctrl', 'v')
                            else:
                                # Или просто вставляем, если не совпадает (на всякий случай)
                                clear_text_field()
                                to_paste = current_values[field_idx]
                                pyperclip.copy(to_paste)
                                time.sleep(0.05)
                                pyautogui.hotkey('ctrl', 'v')

                        text_index += 1

                    pyperclip.copy(old_clip)

            # Обновляем previous_values на текущие для следующей итерации
            previous_values = current_values[:]

            print("✅ Готово — пауза 4 сек")
            time.sleep(4)

            iteration += 1

        print("\n🔄 Повторяем по кругу...")

    print("\n🎉 ВСЁ ГОТОВО")  # Не достигнется

# ========= ХОТКЕИ =========
keyboard.add_hotkey('f9', start_recording)
keyboard.add_hotkey('f10', stop_recording)
keyboard.add_hotkey('f11', run_automation)

print("\n===================================================")
print("🤖 АВТОЗАПОЛНИТЕЛЬ ГОТОВ")
print("F9 — запись | F10 — стоп | F11 — запуск")
print("Во время записи кликайте по полям в любом порядке, код определит по плейсхолдерам.")
print("Автоматизация будет повторяться по кругу бесконечно, заменяя предыдущие значения на новые.")
print("Браузер открывается один раз и не закрывается.")
print("===================================================\n")

keyboard.wait()