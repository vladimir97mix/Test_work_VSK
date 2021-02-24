from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

gChromeExeFullPath = r'Resources\GoogleChromePortable\App\Chrome-bin\chrome.exe'
gExtensionFullPathList = []
gWebDriverFullPath = r'Resources\SeleniumWebDrivers\Chrome\chromedriver_win32 v84.0.4147.30\chromedriver.exe'


def WebDriverInit(inWebDriverFullPath, inChromeExeFullPath, inExtensionFullPathList):
    # Set full path to exe of the chrome
    lWebDriverChromeOptionsInstance = webdriver.ChromeOptions()
    lWebDriverChromeOptionsInstance.binary_location = inChromeExeFullPath
    # Add extensions
    for lExtensionItemFullPath in inExtensionFullPathList:
        lWebDriverChromeOptionsInstance.add_extension(lExtensionItemFullPath)
    # Run chrome instance
    lWebDriverInstance = None
    if inWebDriverFullPath:
        # Run with specified web driver path
        lWebDriverInstance = webdriver.Chrome(executable_path=inWebDriverFullPath,
                                              options=lWebDriverChromeOptionsInstance)
    else:
        lWebDriverInstance = webdriver.Chrome(options=lWebDriverChromeOptionsInstance)
    # Return the result
    return lWebDriverInstance


def FindElemets(webDriver):
    # Переходим на нужную страницу
    webDriver.get('https://yandex.ru/')
    # Таймаут ожидания
    wait = WebDriverWait(webDriver, 10)
    # Дожидаемся до возможности кликнуть на строку поиска
    element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input.input__control.input__input")))
    # Отправляем запрос в поисковую строку
    element.send_keys("Купить страховку")
    # Нажимаем на ENTER
    element.send_keys(Keys.RETURN)



inWebDriver = WebDriverInit(gWebDriverFullPath, gChromeExeFullPath, gExtensionFullPathList)

FindElemets(inWebDriver)
