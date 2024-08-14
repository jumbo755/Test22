from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.keys import Keys
import openpyxl
import time
import random
import string
from faker import Faker
from selenium.common.exceptions import TimeoutException


fake = Faker()
workbook = openpyxl.load_workbook('Test.xlsx')
sheet = workbook.active
password = sheet['C2'].value
login = sheet['B2'].value

EdgeChromiumDriverManager().install() # скачивает и устанавливает драйвер
# Создаем экземпляр WebDriver для Edge, автоматически загружая необходимый драйвер
driver = webdriver.Edge()
driver.get("https://mail.google.com/mail/u/0/#inbox")
wait = WebDriverWait(driver, 5) # Ждем 10 секунд, перед появлением опр элементов, driver: Это объект WebDriver, который управляет браузером (в данном случае Edge)
search_box = wait.until(EC.presence_of_element_located((By.NAME, "identifier")))
search_box.send_keys(login)
search_box.send_keys(Keys.ENTER)

try:
    WebDriverWait(driver, 7).until(
        EC.presence_of_element_located((By.ID, "Passwd"))
    )
    print ("Элемент найден!")
except:
    print("Пароль входа успешно введен")

search_box = wait.until(EC.presence_of_element_located((By.NAME, "Passwd")))
search_box.send_keys(password)
search_box.send_keys(Keys.ENTER)

time.sleep(10)

#Ниже сохраняет резервную почту

driver.get("https://myaccount.google.com/recovery/email?continue=https://myaccount.google.com/security?hl%3Den%26rapt%3DAEjHL4NIHMgM7YJk8MTdFos8oJepnL2EVL48HyiW8JSB3evhnKkvRahKClkmv_nqgoshS2I6ruXKgauv9GQRFgPc1VSvJK9z0fad2GgZoY1mkabwNny-J8c%26utm_source%3DOGB%26utm_medium%3Dact&rapt=AEjHL4OZuXLHisTgtC6_ohcqjjQ-dZcQNip-PcrYYA6hMGRgNKofT6wi0INvW4JEzXuKAwv8U0o9mP0_dgMDf-z4fmcv3XLU7ElfXzdFv7iPFsCy1k2w3R0")
element = driver.find_element(By.XPATH, "//input[@id='i5']")
value = element.get_attribute("value")
sheet['F2'] = value
workbook.save('Test.xlsx')

time.sleep(10)

#Ниже сохраняет дату рождения
driver.get("https://myaccount.google.com/personal-info?hl=en&authuser=0&rapt=AEjHL4NIHMgM7YJk8MTdFos8oJepnL2EVL48HyiW8JSB3evhnKkvRahKClkmv_nqgoshS2I6ruXKgauv9GQRFgPc1VSvJK9z0fad2GgZoY1mkabwNny-J8c&utm_source=OGB&utm_medium=act")
try:
    # Находим ссылку по href (приспособьте XPath под свои нужды)
    link = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//a[contains(@href, 'birthday')]"))
    )

    # Получаем значение aria-label
    aria_label = link.get_attribute("aria-label")

    sheet['E2'] = aria_label

    print("Aria-label:", aria_label)
    workbook.save('Test.xlsx')

except TimeoutException:
    print("Элемент не найден")
except Exception as e:
    print("Произошла ошибка:", e)

time.sleep(10)

#Ниже меняет и сохраняет имя и фамилию аккаунта

name = fake.first_name()
name1 = fake.last_name()
sheet['H2'] = name
sheet['I2'] = name1
workbook.save('Test.xlsx')

driver.get("https://myaccount.google.com/profile/name/edit?continue=https://myaccount.google.com/profile/name?continue%3Dhttps%253A%252F%252Fmyaccount.google.com%252Fpersonal-info%253Fhl%253Den%2526utm_source%253DOGB%2526utm_medium%253Dact%2526rapt%253DAEjHL4NIHMgM7YJk8MTdFos8oJepnL2EVL48HyiW8JSB3evhnKkvRahKClkmv_nqgoshS2I6ruXKgauv9GQRFgPc1VSvJK9z0fad2GgZoY1mkabwNny-J8c%2526pli%253D1%26hl%3Den%26rapt%3DAEjHL4NIHMgM7YJk8MTdFos8oJepnL2EVL48HyiW8JSB3evhnKkvRahKClkmv_nqgoshS2I6ruXKgauv9GQRFgPc1VSvJK9z0fad2GgZoY1mkabwNny-J8c%26utm_source%3DOGB%26utm_medium%3Dact&pli=1&rapt=AEjHL4Nfy6jUTdY3a3zZDmjTncAlrPakASK1z836OwH4Bcs96divlRkKDmC9DWjZB-ICA3QlhVX1d8nDQ_rUmxO3OxXoUp9LWwFYBBF3Gxn56ya1foQMh7Y")
wait = WebDriverWait(driver, 5) # Ждем 10 секунд, перед появлением опр элементов, driver: Это объект WebDriver, который управляет браузером (в данном случае Edge)
element = driver.find_element(By.CSS_SELECTOR, "[jsname='vhZMvf']").click()
driver.find_element(By.CSS_SELECTOR, "[jsname='vhZMvf']").send_keys(Keys.CONTROL + "a")
driver.find_element(By.CSS_SELECTOR, "[jsname='vhZMvf']").send_keys(name)
driver.find_element(By.CSS_SELECTOR, "[jsname='vhZMvf']").send_keys(Keys.TAB + name1)
time.sleep(3)
driver.find_element(By.CSS_SELECTOR, "[jsname='vhZMvf']").send_keys(Keys.TAB + Keys.TAB + Keys.TAB + Keys.TAB + Keys.ENTER )

time.sleep(10)


#Ниже меняет и сохраняет пароль
def generate_password(length=12):
    """Генерирует случайный пароль заданной длины.

    Args:
        length: Длина пароля.

    Returns:
        str: Сгенерированный пароль.
    """

    # Все возможные символы для пароля
    characters = string.ascii_letters + string.digits + string.punctuation

    password = ''.join(random.choice(characters) for _ in range(length))
    return password
password = generate_password()

print(password)
sheet['C2'] = password
workbook.save('Test.xlsx')
print("Пароль успешно сохранен в ячейку G1")


driver.get("https://accounts.google.com/v3/signin/challenge/pwd?TL=AKeb6mzEDv-aUt48B3IIWtaWrRdBfyZhjNc8gB8yaNUwiT-CTrl58AUaR2zyLlLC&cid=1&continue=https%3A%2F%2Fmyaccount.google.com%2Fsigninoptions%2Fpassword%3Fcontinue%3Dhttps%3A%2F%2Fmyaccount.google.com%2Fsecurity%3Fhl%253Den%2526utm_source%253DOGB%2526utm_medium%253Dact&flowName=GlifWebSignIn&hl=en&ifkv=Ab5oB3roRd4dC2t2cgNzLXsT3jYl1wAUP92Pj0fje3itAT1HMUzl6knwDfbwb1pJ1M6LPtAmk6So7A&kdi=CAM&rart=ANgoxcdUkMnMJRVK-D-0ptfNJHFWcZIVAQkTHZaVPjJ1s2HQ93O6yYOT5NQpqhdgol_j1WGrWY--4x1qjhF7m2X6M0lWWDI7EwU-15rM-J3NLIsSkJBEsDA&rpbg=1&sarp=1&scc=1&service=accountsettings")

try:
    WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.ID, "password"))
    )
    print ("Элемент найден!")
except:
    print("Смена пароля 1 успешна введена")


search_box = wait.until(EC.presence_of_element_located((By.NAME, "password")))
search_box.send_keys(password)
search_box.send_keys(Keys.ENTER)

try:
    WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.ID, "confirmation_password"))
    )
    print ("Элемент найден!")
except:
    print("Смена пароля 2 успешна введена")


search_box = wait.until(EC.presence_of_element_located((By.NAME, "confirmation_password")))
search_box.send_keys(password)
search_box.send_keys(Keys.ENTER)

button = driver.find_element(By.ID, "Pr7Yme")
button.click()

driver.quit()
