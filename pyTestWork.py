import os
import time
import shutil
from docx import Document
from docx.shared import Inches
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

gChromeExeFullPath = r'Resources\GoogleChromePortable\App\Chrome-bin\chrome.exe'
gExtensionFullPathList = []
gWebDriverFullPath = r'Resources\SeleniumWebDrivers\Chrome\chromedriver_win32 v84.0.4147.30\chromedriver.exe'


def WebDriverInit(inWebDriverFullPath, inChromeExeFullPath, inExtensionFullPathList):
    # Определение полного пути к брузеру и драйверу
    lWebDriverChromeOptionsInstance = webdriver.ChromeOptions()
    lWebDriverChromeOptionsInstance.binary_location = inChromeExeFullPath
    # Добавление расширений
    for lExtensionItemFullPath in inExtensionFullPathList:
        lWebDriverChromeOptionsInstance.add_extension(lExtensionItemFullPath)
    # Запустить экземпляр Chrome
    lWebDriverInstance = None
    if inWebDriverFullPath:
        # Запустить с указанным путем веб-драйвера
        lWebDriverInstance = webdriver.Chrome(executable_path=inWebDriverFullPath,
                                              options=lWebDriverChromeOptionsInstance)
    else:
        lWebDriverInstance = webdriver.Chrome(options=lWebDriverChromeOptionsInstance)
    # Вернуть результат
    return lWebDriverInstance


def FindElemets(webDriver):
    webDriver.get('https://yandex.ru/')  # Переходим на нужную страницу
    wait = WebDriverWait(webDriver, 10)  # Таймаут ожидания
    # Дожидаемся до возможности кликнуть на строку поиска
    element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input.input__control.input__input")))
    element.send_keys("Купить страховку")  # Отправляем запрос в поисковую строку
    element.send_keys(Keys.RETURN)  # Нажимаем на ENTER

    elementMain = webDriver.find_elements(By.CLASS_NAME, 'serp-item')  # Основные эелементы поиска (li)
    elementLink = webDriver.find_elements(By.CLASS_NAME, 'OrganicTitle-Link')  # Линки для перехода по ссылкам

    titleTextList = []  # Список заголовков
    extendetTextLis = []  # Список подзаголовков

    for elem in elementMain:
        # Ищем в элементах заголовки и вносим текст в список
        if elem.find_elements(By.CLASS_NAME, 'OrganicTitle-LinkText'):
            elem_text = elem.find_elements(By.CLASS_NAME, 'OrganicTitle-LinkText')
            titleTextList.append(elem_text[0].text)
        # Ищем в элементах позаголовки с классом Organic-Text и вносим текст в список
        if elem.find_elements(By.CLASS_NAME, 'Organic-Text'):
            elem_text = elem.find_elements(By.CLASS_NAME, 'Organic-Text')
            extendetTextLis.append(elem_text[0].text)
        # Ищем в элементах позаголовки с классом text-container и вносим текст в список
        if elem.find_elements(By.CLASS_NAME, 'text-container.typo.typo_text_m.typo_line_m.organic__text'):
            elem_text = elem.find_elements(By.CLASS_NAME, 'text-container.typo.typo_text_m.typo_line_m.organic__text')
            extendetTextLis.append(elem_text[0].text)

    # Создаем папку для скринов
    if not os.path.exists('screens'):
        os.makedirs('screens')
    picNum = 1
    # Открытие ссылок результатов
    for elem in elementLink:
        elem.click()
        handles = webDriver.window_handles  # Список открытых вкладок в браузере


        for tab in handles[1:]:
            webDriver.switch_to_window(tab)  # Переключаемся по каждой вкладке, кроме первой - yandex.ru
            webDriver.get_screenshot_as_file('screens/scr' + str(picNum) + '.png')  # Делаем скрин в папку 'screens'
            picNum += 1
            webDriver.close()
            webDriver.switch_to_window(handles[0])


    picNum -= 1
    strNum = 1  # Нумерация заголовка
    doc = Document()
    for txtTitle, txtExt in zip (titleTextList, extendetTextLis):  # Циклл по спискам заголовков
        p = doc.add_paragraph()  # Создаем новый абзац
        run = p.add_run(str(strNum) + ' ' + txtTitle)  # Добавляем текст
        run.bold = True  # Делаем текст жирным
        doc.add_paragraph(txtExt)  # Вставляем подзаголовок в новый абзац
        doc.add_picture('screens/scr' + str(picNum) + '.png', width=Inches(4.0))  # Вставляем скриншот, шириной 4 дюйма
        picNum -= 1  # Декримент скринов
        strNum += 1  # Инкримент номера заголовка

    doc.save('Отчет по ключевой фразе_{}.docx'.format(int(time.time())))  # Сохранение файла

    shutil.rmtree('screens/', ignore_errors=True)  # Удаление папки скринов за ненадобностью

    webDriver.quit()  # Закрытие браузера

# Инициализация веб драйвера
inWebDriver = WebDriverInit(gWebDriverFullPath, gChromeExeFullPath, gExtensionFullPathList)

# Запуск задачи для веб драйвера
FindElemets(inWebDriver)


