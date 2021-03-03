import os
import time
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from pyOpenRPA.Robot import UIDesktop

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


titleTextList = []  # Список заголовков
extendetTextLis = []  # Список подзаголовков


def FindElemets(webDriver):
    webDriver.get('https://yandex.ru/')  # Переходим на нужную страницу
    wait = WebDriverWait(webDriver, 10)  # Таймаут ожидания
    # Дожидаемся до возможности кликнуть на строку поиска
    element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input.input__control.input__input")))
    element.send_keys("Купить страховку")  # Отправляем запрос в поисковую строку
    element.send_keys(Keys.RETURN)  # Нажимаем на ENTER

    elementMain = webDriver.find_elements(By.CLASS_NAME, 'serp-item')  # Основные эелементы поиска (li)
    elementLink = webDriver.find_elements(By.CLASS_NAME, 'OrganicTitle-Link')  # Линки для перехода по ссылкам

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
    picNum = 1  # Нумерация скриншотов
    # Открытие ссылок результатов
    for elem in elementLink:
        elem.click()
        handles = webDriver.window_handles  # Список открытых вкладок в браузере

        for tab in handles[1:]:
            webDriver.switch_to_window(tab)  # Переключаемся по каждой вкладке, кроме первой - yandex.ru
            webDriver.get_screenshot_as_file('screens/scr' + str(picNum) + '.png')  # Делаем скрин в папку 'screens'
            picNum += 1  # Инкримент номера скринов
            webDriver.close()  # Закрытие текущей вкладки
            webDriver.switch_to_window(handles[0])  # Переход на основную вкладку

    # Запись текста эелементов страницы в wordpad
    wordPadWriter()


def wordPadWriter():
    picNum = 1  # Нумерация скриншотов
    strNum = 1  # Нумерация заголовка

    # uio селекторы:
    wpSel = [
        {"title": "Документ - WordPad", "class_name": "WordPadClass",
         "backend": "win32"}]  # Сформировали UIO селектор
    wpRt = [{"title": "Документ - WordPad", "class_name": "WordPadClass", "backend": "win32"},
            {"title": "Окно Rich Text", "rich_text": "", "ctrl_index": 4}]
    wpPic = [{"class_name": "WordPadClass", "backend": "uia"},
             {"title": "Изображение", "depth_start": 11, "depth_end": 11}]
    wpFn = [{"class_name": "WordPadClass", "backend": "uia"},
            {"title": "Имя файла:", "depth_start": 3, "depth_end": 3}]
    wpFn2 = [{"class_name": "WordPadClass", "backend": "uia"},
             {"title": "Имя файла:", "depth_start": 4, "depth_end": 4}]

    lExistBool = UIDesktop.UIOSelector_Exist_Bool(inUIOSelector=wpSel)  # Проверить наличие окна по UIO селектору
    if not lExistBool:  # Проверить наличие окна wordpad
        os.system("write")  # Открыть wordpad
        time.sleep(2)
    else:  # Проверить, что окно wordpad не свернуто
        UIOWordPad = UIDesktop.UIOSelector_Get_UIO(inSpecificationList=wpSel)  # Получить UIO экземпляр
        if UIOWordPad.is_minimized():  # Проверить, что wordpad находится в свернутом виде
            UIOWordPad.restore()  # Восстановить окно wordpad из свернутого вида
        # Проверить наличие UI элемента по UIO селектору
    lWP_IsExistBool = UIDesktop.UIOSelector_Exist_Bool(inUIOSelector=wpRt)
    if lWP_IsExistBool:
        uioRichText = UIDesktop.UIOSelector_Get_UIO(inSpecificationList=wpRt)
        uioPasteImg = UIDesktop.UIOSelector_Get_UIO(inSpecificationList=wpPic)

    # Ввод текста из списков, в wordpad
    for txtTitle, txtExtTitle in zip(titleTextList, extendetTextLis):
        uioRichText.type_keys('^b''^и')
        uioRichText.type_keys(str(strNum) + '{SPACE}' + txtTitle, with_spaces=True)
        uioRichText.type_keys('{ENTER}')
        uioRichText.type_keys('^b''^и')
        uioRichText.type_keys(txtExtTitle, with_spaces=True)
        uioRichText.type_keys('{ENTER}')

        # Вставка скриншотов
        uioPasteImg.click_input()
        if UIDesktop.UIOSelector_Exist_Bool(wpFn):
            uioFilename = UIDesktop.UIOSelector_Get_UIO(wpFn)
        uioFilename.type_keys(os.path.abspath('screens/scr' + str(picNum) + '.png'))
        uioFilename.type_keys('{ENTER}')
        uioRichText.type_keys('{ENTER}')
        strNum += 1  # Инкримент номера заголовка
        picNum += 1  # Инкримент номера скрина

    # Сохраниение файла
    uioRichText.type_keys('^s''^ы')  # Горячая клавиша "сохранить документ"
    if UIDesktop.UIOSelector_Exist_Bool(wpFn2):
        uioFilename = UIDesktop.UIOSelector_Get_UIO(wpFn2)  # Выбор строки ввода пути к файлу
        # Ввод пути к файлу
        uioFilename.type_keys(os.path.abspath('Отчет по ключевой фразе_{}.rtf'.format(int(time.time()))),
                              with_spaces=True)
        uioFilename.type_keys('{ENTER}')



# Инициализация веб драйвера
inWebDriver = WebDriverInit(gWebDriverFullPath, gChromeExeFullPath, gExtensionFullPathList)

# Поиск необходимых элементов
FindElemets(inWebDriver)

shutil.rmtree('screens/', ignore_errors=True)  # Удаление папки скринов за ненадобностью
