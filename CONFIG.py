import random
import time
import os
import shutil
import pyautogui
import mouseinfo

# debug функция
def debug(debug_text: str):
    with open("src/log.txt", "a") as file:
        file.write("\nDEBUG: " + debug_text)
    print("DEBUG: " + debug_text)

# Выбор случайного user-agent'а из списка
def get_random_user_agent():
    with open('src/user-agents.txt', 'r') as f:
        lines = f.readlines()
    return random.choice(lines).strip()

# ожидание от 1 до 3 секунд
def wait_random():
    wait_time = random.randint(0, 1)
    time.sleep(wait_time)

def backup_file(count, path_to_file):
    # Создаем папку "backups", если ее еще нет
    if not os.path.exists('backups'):
        os.makedirs('backups')

    # Определяем имя файла и его путь
    filename = path_to_file
    basename = os.path.basename(filename)
    backup_name = f'backup{count}.xlsx'
    backup_path = os.path.join('backups', backup_name)

    # Копируем файл в папку "backups" с новым именем
    shutil.copyfile(filename, backup_path)
    debug(f'Файл {count} успешно скопирован в ' + backup_path)

def read_progress():
    with open("src/progress.txt", "r") as f:
        progress = int(f.read().strip())
    return progress

def write_progress(progress):
    with open("src/progress.txt", "w") as f:
        f.write(str(progress))

def clicker_vpn():
    global progress_bar

    # Запуск кликера
    progress_bar -= 2
    debug(f"Запуск кликера & PROGRESS-2 {progress_bar}")
    time.sleep(2)

    # Двигаем мышку в верхний правый угол экрана на x1 пикселей
    pyautogui.moveTo(pyautogui.size().width - 290, 20)
    print("Двигаем мышку в верхний правый угол экрана на x пикселей")
    time.sleep(2)

    # Нажимаем ЛКМ
    pyautogui.click()
    print("Нажимаем ЛКМ на впн")

    # Двигаем мышку вниз на x2 пикселей
    pyautogui.moveRel(0, 310)
    print("Двигаем мышку вниз на x пикселей")
    time.sleep(2)

    # Нажимаем ЛКМ
    pyautogui.click()
    print("Нажимаем ЛКМ на Регион")

    # Двигаем мышку вниз на x3 пикселей
    pyautogui.moveRel(0, 60)
    print("Двигаем мышку вниз на x3 пикселей")
    time.sleep(2)

    # Крутим колесо мышки 10 раз
    rand_scroll = random.randint(1, 80)
    pyautogui.scroll(-rand_scroll)
    print(f"Крутим колесо мышки {rand_scroll} раз")
    time.sleep(2)

    # Нажимаем ЛКМ
    pyautogui.click()

    # Ждем 10 секунд
    print("Ждем 8 секунд")
    time.sleep(2)

    # Перемещаем мышку на кнопку обновить
    pyautogui.moveRel(-1452, -280)
    print("Перемещаем мышку на кнопку обновить")
    time.sleep(2)

    # Обновляем браузер
    pyautogui.click()
    print("Обновляем браузер")
    time.sleep(1)

    # Выводим сообщение
    print("Кликер сработал")

# окведы
OKVEDS_TRUE = ['62.03.12', '62.03.13', '62.03.19', '62.03.11', '63.99.11', '63.99.12',
               '62.02.1', '62.02.2', '62.02.3', '62.02.4', '62.02.9', '62.03.1', '63.12.1',
               '63.11.1', '63.11.9', '63.99.1', '63.99.2', '63.91', '63.99', '63.11', '63.12',
               '62.01', '62.02', '62.03', '62.09', '63.9', '63.0', '62.0']

# ссылка на сайт с поиском
path_to_file = "src/company.xlsx"
url = "https://www.rusprofile.ru/search?query="
copy_const = 100

# счетчики
progress_bar = read_progress()
company_name = ''
company_index = ''