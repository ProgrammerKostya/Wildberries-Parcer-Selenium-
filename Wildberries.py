import time

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import xlsxwriter

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("--disable-cache")
chrome_options.add_argument("--window-size=1920,1080") # Если нужно врубить headless, закомментируй
chrome_options.add_argument('--incognito')

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
action = ActionChains(driver)
wildberries_url = "https://www.wildberries.ru/"
wait = WebDriverWait(driver, 10, 1)

total_page = 0 # Отметка о странице, на которой мы находимся!!
fixed_step = 0 # Количество итераций, которые мы должны совершить!!
count = 2 # Счетчик итераций в Excel

"""
Создание Excel документа!!
"""
workbook = xlsxwriter.Workbook('wildberries.xlsx')
worksheet = workbook.add_worksheet()


def Excel(data, start):
    global count
    # Заголовки столбцов
    headers = ['Товар', 'Цена', 'Рейтинг', 'Когда будет доставлен', 'Ссылка']
    worksheet.write_row('A1', headers)
    # Данные для таблицы
    data = data

    # Записываем данные в Excel
    for row_num, row_data in enumerate(data, start=start):  # Начинаем с 2, чтобы пропустить заголовки
        worksheet.write_row(row_num, 0, row_data)
        count += 1

def repeater(steps):
    for step in range(1, steps + 1):
        yield step

def collector(rates, urls, names, goods, expected_price, dates):
    global total_page
    global count
    data = []
    start_position = count
    for good in goods:
        new_list = []
        new_good = good.text
        new_good = new_good.replace("₽", '').replace(' ', '')
        new_good = int(new_good) # Цена
        if new_good <= expected_price:
            index = goods.index(good)
            new_list.append(names[index].text.replace('/', '').strip()) # Наименование товара
            new_list.append(new_good) # Цена на товар
            new_list.append(rates[index].text) # Рейтинг на товар
            new_list.append(dates[index].text) # Дата
            new_list.append(urls[index].get_attribute("href")) # Ссылка
            data.append(new_list)
    Excel(data, start=start_position)
    total_page += 1

def body(price, number):
    global fixed_step
    global total_page
    expected_price = price
    number_page = number
    time.sleep(5)

    while True:
        element = wait.until(EC.presence_of_element_located(("xpath", "//div[contains(@class, 'pageToInsert')]")))
        last_scroll_position = driver.execute_script("return window.scrollY;")
        action.scroll_to_element(element).perform()
        current_scroll_position = driver.execute_script("return window.scrollY;")
        if last_scroll_position == current_scroll_position:
            driver.execute_script("""
                window.scrollTo({
                top: window.scrollY + 100,
                });
                """)
            break

    time.sleep(5)
    goods = driver.find_elements("xpath", "//del")
    names = driver.find_elements("xpath", "//span[@class='product-card__name']")
    urls = driver.find_elements("xpath", "//a[contains(@class, 'product-card__link')]")
    rates = driver.find_elements("xpath", "//span[contains(@class, 'address-rate-mini')]")
    dates = driver.find_elements("xpath", "//span[@class='btn-text']")
    collector(urls=urls, names=names, goods=goods, expected_price=expected_price, rates=rates, dates=dates)
    if number_page >= fixed_step:
        print(f"Программа завершила свою работу и спарсила {total_page} страниц!!")
        # Закрываем файл
        workbook.close()

        print("Excel файл успешно создан!")
        driver.quit()
        driver.close()
    else:
        try:
            time.sleep(5)
            driver.find_element("xpath", f"//a[text()='{number_page + 1}']").click()
        except Exception:
            pass


def main():
    global fixed_step
    try:
        item = input("Какой товар вас интересует?: ")
        price = float(input("Какая цена вас интересует?: "))
        numbers = int(input("Сколько страниц нужно спарсить?: "))
        driver.get(wildberries_url)

        time.sleep(5)

        SEARCH = ('xpath', "//input[@id='searchInput']")
        search_field = wait.until(EC.element_to_be_clickable(SEARCH))
        search_field.send_keys(item)
        time.sleep(3)
        search_field.send_keys(Keys.ENTER)

        time.sleep(5)

        fixed_step = numbers
        repeater(numbers)
        for step in repeater(numbers):
            body(price=price, number=step)
    except ValueError:
        print("Введите price и numbers цифрами!!")
        print("Кол-во страниц должно быть ЦЕЛЫМ")

if __name__ == "__main__":
    main()
else:
    print("Программа должна открываться со своей страницы!!")