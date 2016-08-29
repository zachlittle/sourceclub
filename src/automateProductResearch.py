from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

path = 'C:/Users/PC/programming/sourceclub/env/selenium/webdriver/chrome/chromedriver.exe'
driver = webdriver.Chrome(path)

LOGINCREDENTIALS = {'Username': 'sourceclubcommerce@gmail.com',
                    'Password': 'chronicboys760'}


class LOWPRICEPRODUCT:
    CATEGORIES = ['Home Improvement', 'Automotive', 'Clothing', 'Office Products']
    PRICERANGE = ['10', '30']
    NET = ['8']

class MEDIUMPRICEPRODUCT:
    CATEGORIES = ['Home Improvement', 'Automotive', 'Clothing', 'Office Products']
    PRICERANGE = ['31', '50']
    NET = ['8']

class HIGHPRICEPRODUCT:
    CATEGORIES = ['Home Improvement', 'Automotive', 'Clothing', 'Office Products']
    PRICERANGE = ['51', '80']
    NET = ['8']

class AutomateJSWork:

    def __init__(self):
        AutomateJSWork.Login(self)
        AutomateJSWork.InputSearchTypeData(LOWPRICEPRODUCT)
        AutomateJSWork.InputSearchTypeData(MEDIUMPRICEPRODUCT)
        AutomateJSWork.InputSearchTypeData(HIGHPRICEPRODUCT)

    def Login(self):
        driver.get('https://www.junglescout.com/')
        driver.implicitly_wait(100)
        driver.find_element_by_id('menu-item-4544').click()
        inputElementUsername = driver.find_element_by_id('user_login')
        inputElementUsername.send_keys(LOGINCREDENTIALS['Username'])
        inputElementPassword = driver.find_element_by_id('user_password')
        inputElementPassword.send_keys(LOGINCREDENTIALS['Password'])
        driver.find_element_by_name('commit').click()

    def InputSearchTypeData(searchType):
        print("CURRENTLY RUNNING " + str(searchType) + " SEARCH")
        driver.implicitly_wait(10)
        driver.find_element_by_class_name('database-nav').click()
        AutomateJSWork.clickDesiredCategories(searchType)
        AutomateJSWork.enterPriceRange(searchType)
        AutomateJSWork.enterNetRange(searchType)
        driver.quit()

    def clickDesiredCategories(searchType):
        driver.refresh()
        necItemsToClick = len(searchType.CATEGORIES)
        if necItemsToClick > 0:
            for listItem in searchType.CATEGORIES:
                if necItemsToClick > 0:
                    print(listItem)
                    allElem = driver.find_elements_by_xpath(".//span[@class = 'category-label']")
                    print(allElem)
                    for item in allElem:
                        if listItem in item.text:
                            item.click()
                            necItemsToClick -= 1
                else:
                    break

    def enterPriceRange(searchType):
        allElem = driver.find_elements_by_xpath(".//span[@class = 'db-filter-content']")
        for item in allElem:
            if "Price" in item.text:
                if searchType.PRICERANGE[0] is not None:
                    inputMinBox = item.find_element_by_xpath("//input[@placeholder = 'min']")
                    inputMinBox.send_keys(Keys.CONTROL + "a")
                    inputMinBox.send_keys(Keys.DELETE)
                    inputValMin = str(searchType.PRICERANGE[0])
                    inputMinBox.send_keys(inputValMin)
                if len(searchType.PRICERANGE) >= 2:
                    inputMaxBox = item.find_element_by_xpath("//input[@placeholder = 'max']")
                    inputMaxBox.send_keys(Keys.CONTROL + "a")
                    inputMaxBox.send_keys(Keys.DELETE)
                    inputValMax = str(searchType.PRICERANGE[1])
                    inputMaxBox.send_keys(inputValMax)

    def enterNetRange(searchType):
        allElem = driver.find_elements_by_xpath(".//span[@class = 'db-filter-content']")
        for item in allElem:
            print(item.text)
            if "Net" in item.text:
                if searchType.NET[0] is not None:
                    inputMinBox = item.find_elements_by_xpath("//input[@placeholder = 'min']")
                    print(inputMinBox)
                    inputMinBox[1].send_keys(Keys.CONTROL + "a")
                    inputMinBox[1].send_keys(Keys.DELETE)
                    inputValMin = str(searchType.NET[0])
                    inputMinBox[1].send_keys(inputValMin)
                if len(searchType.NET) >= 2:
                    inputMaxBox = item.find_elements_by_xpath("//input[@placeholder = 'max']")
                    print(inputMinBox)
                    inputMinBox[1].send_keys(Keys.CONTROL + "a")
                    inputMaxBox[1].send_keys(Keys.DELETE)
                    inputValMax = str(searchType.NET[1])
                    inputMaxBox[1].send_keys(inputValMax)


AutomateJSWork()