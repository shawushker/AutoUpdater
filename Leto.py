import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import psutil

# Путь к вашему файлу
file_path = 'C:/Users/AUTOMATIZ/YandexDisk/Update/products-1415348.xlsx'

# Путь к пользовательским данным Chrome
user_data_dir = 'C:/Users/AUTOMATIZ/AppData/Local/Google/Chrome/User Data'
profile_directory = 'Default'  # Используйте 'Default' для профиля по умолчанию или замените на нужный профиль

# Проверка существования пути к пользовательским данным
if not os.path.exists(user_data_dir):
    raise Exception(f"User data directory does not exist: {user_data_dir}")

# Настройка опций для Chrome
options = uc.ChromeOptions()
options.add_argument('--no-first-run')
options.add_argument('--no-service-autorun')
options.add_argument('--password-store=basic')
options.add_argument(f'--user-data-dir={user_data_dir}')
options.add_argument(f'--profile-directory={profile_directory}')
options.add_argument('--lang=ru')  # Установка языка на русский
options.add_argument('--force-encoding-to-utf8')  # Принудительная установка кодировки UTF-8

try:
    # Запуск драйвера
    driver = uc.Chrome(version_main=127, options=options)
    print("Driver started successfully.")

    # Переход на нужную страницу после установки куки
    driver.get('https://seller.ozon.ru/app/highlights/my-highlights/1415348?page=1&pageSize=20')
    print("Navigated to the page.")
    time.sleep(5)  # Дайте время странице загрузиться

    # Нажатие на кнопку "Добавить товары" с использованием JavaScript
    add_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable(
            (By.CSS_SELECTOR, '.custom-button_button_5t4mr.custom-button__theme_secondary_3xMsy.custom-button__size_normal_xYW59.custom-button_textAlignCenter_13fpm.header_button_1J-CE')
        )
    )
    driver.execute_script("arguments[0].click();", add_button)
    print("Clicked 'Add Products' button.")
    time.sleep(2)

    # Попытка найти кнопку "Импорт в XLS-файле" с помощью различных селекторов и с использованием JavaScript
    try:
        import_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Импорт в XLS-файле')]"))
        )
        print("Found 'Import in XLS file' button by XPath.")
    except:
        import_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, '.custom-button_button_5t4mr.custom-button__theme_primary_1oBem.custom-button__size_normal_xYW59.custom-button_textAlignCenter_13fpm')
            )
        )
        print("Found 'Import in XLS file' button by CSS Selector.")
    
    driver.execute_script("arguments[0].click();", import_button)
    print("Clicked 'Import in XLS file' button.")
    time.sleep(2)

    # Нажатие на кнопку "Выбрать файл" с использованием JavaScript
    choose_file_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Выбрать файл')]"))
    )
    driver.execute_script("arguments[0].click();", choose_file_button)
    print("Clicked 'Choose File' button.")
    time.sleep(2)

    # Выбор файла и загрузка
    file_input = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, '//input[@type="file"]'))
    )
    file_input.send_keys(file_path)
    print("File path entered.")
    time.sleep(20)

    # Нажатие на кнопку "Сохранить изменения" с использованием JavaScript
    save_button = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Сохранить изменения')]/ancestor::button"))
    )
    driver.execute_script("arguments[0].click();", save_button)
    print("Clicked 'Save Changes' button.")
    time.sleep(30)

except Exception as e:
    print(f"An error occurred: {e}")
finally:
    # Закрытие драйвера
    driver.quit()
    print("Driver closed.")

    # Завершение всех процессов Chrome и драйвера
    for process in psutil.process_iter():
        if process.name() in ["chrome.exe", "chromedriver.exe"]:
            process.kill()
    print("All Chrome and driver processes terminated.")
