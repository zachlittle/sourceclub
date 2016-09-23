from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import src.excelParser


path = 'C:/Git/sourceclub/src/chromedriver.exe'
driver = webdriver.Chrome(path)

LOGINCREDENTIALS = {'Username': 'sourceclubcommerce@gmail.com',
                    'Password': 'chronicboys760'}

MINMAXINDEXMAP = {'Price': 0,
                  'Net': 1,
                  'Rank': 2,
                  'Est. Sales': 3,
                  'Est. Rev': 4,
                  'Reviews': 5,
                  'Rating': 6,
                  'Weight': 7,
                  'No. Sellers': 8,
                  'Listing Quality': 9}

MEDIUMPRICEPRODUCT = {
    'Weight': [0, 10],
    'Price': [20, 100],
    'Rating': [0, 0],
    'Rank': [0, 0],
    'Reviews': [0, 200],
    'Net': [10, 0],
    'Listing Quality': [5, 90],
    'No. Sellers': [0, 15],
    'Est. Rev': [0, 0],
    'Est. Sales': [0, 300],
    'Categories': ['Home Improvement', 'Automotive', 'Clothing', 'Office Products'],
    'ProductTier': ['Standard', 'Oversize'],
    'Seller': ['Amazon', 'Fulfilled Amazon', 'Fulfilled Merchant'],
    'Excluded Keyword': 'bible',
    'Keyword': 'Goats'
}


class AutomateJSWork:
    def __init__(self):
        #try:
        self.Login()
        self.clickDatabase()
        self.InputSearchTypeData()
        self.Logout()
        driver.quit()
        src.excelParser.excelHandler()
        '''except:
            e = sys.exc_info()[0]
            print('Exception: ' + str(e))'''

    @staticmethod
    def Login():
        driver.get('https://www.junglescout.com/')
        driver.maximize_window()
        driver.implicitly_wait(100)
        driver.find_element_by_id('menu-item-4544').click()
        inputElementUsername = driver.find_element_by_id('user_login')
        inputElementUsername.send_keys(LOGINCREDENTIALS['Username'])
        inputElementPassword = driver.find_element_by_id('user_password')
        inputElementPassword.send_keys(LOGINCREDENTIALS['Password'])
        driver.find_element_by_name('commit').click()

    @staticmethod
    def Logout():
        element = driver.find_element_by_xpath(".//span[@class = 'username username-hide-on-mobile']")
        element.click()
        element2 = driver.find_element_by_xpath(".//a[@href = '/users/sign_out']")
        element2.click()

    @staticmethod
    def clickSearch():
        element = driver.find_element_by_class_name('make-search')
        element.click()


    def InputSearchTypeData(self):
        print("CURRENTLY RUNNING " + str(MEDIUMPRICEPRODUCT) + " SEARCH")
        element = driver.find_element_by_class_name('filter-reset-btn')
        element.click()
        # driver.implicitly_wait(webdriver.Chrome, 10)
        self.waitForLoad('category-aggregations')
        self.inputData(MEDIUMPRICEPRODUCT)
        self.clickResultsPerPage()
        self.clickSearch()
        self.clickExportToExcel()

    def waitForLoad(self, className):
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        Wait = WebDriverWait(driver, 10)
        Wait.until(EC.presence_of_element_located((By.CLASS_NAME, className)))

    @staticmethod
    def clickDatabase():
        driver.find_element_by_class_name('database-nav').click()

    @staticmethod
    def clickResultsPerPage():
        element = driver.find_element_by_xpath("//select[@name = 'select']")
        action = webdriver.common.action_chains.ActionChains(driver)
        action.move_to_element_with_offset(element, 1, 1)
        action.click()
        action.perform()
        element2 = element.find_element_by_xpath("//option[@value = '200']")
        element2.click()

    def inputData(self, searchType):
        self.clickDesiredCategories(searchType)
        minMaxClass = MinMaxInputFieldDriver()
        minMaxClass.enterAllData(searchType)
        self.clickProductTier(searchType)
        self.clickSeller(searchType)
        self.enterSearchKeyword(searchType)
        self.enterExcludedKeyword(searchType)

    def clickDesiredCategories(self, searchType):
        necItemsToClick = len(searchType['Categories'])
        for listItem in searchType['Categories']:
            if necItemsToClick > 0:
                allElem = driver.find_elements_by_xpath(".//span[@class = 'category-label']")
                for item in allElem:
                    if listItem in item.text:
                        item.click()
                        necItemsToClick -= 1
            else:
                break

    def clickProductTier(self, searchType):
        necItemsToClick = len(searchType['ProductTier'])
        for listItem in searchType['ProductTier']:
            if necItemsToClick > 0:
                allElem = driver.find_elements_by_xpath(".//div[@class = 'category-label-content']")
                for item in allElem:
                    if listItem in item.text:
                        item.click()
                        necItemsToClick -= 1
            else:
                break

    def clickSeller(self, searchType):
        necItemsToClick = len(searchType['Seller'])
        for listItem in searchType['Seller']:
            if necItemsToClick > 0:
                allElem = driver.find_elements_by_xpath(".//div[@class = 'category-label-content']")
                for item in allElem:
                    if listItem in item.text:
                        item.click()
                        necItemsToClick -= 1
            else:
                break

    @staticmethod
    def clickSearch():
        element = driver.find_element_by_class_name('make-search')
        element.click()

    @staticmethod
    def clickExportToExcel():
        element = driver.find_element_by_xpath(".//img[@alt = 'download-csv']")
        element.click()

    def enterSearchKeyword(self, searchType):
        element = driver.find_element_by_class_name('db-keyword-search')
        element.send_keys(Keys.CONTROL + "a")
        element.send_keys(Keys.DELETE)
        searchKeyword = str(searchType['Keyword'])
        element.send_keys(searchKeyword)

    def enterExcludedKeyword(self, searchType):
        element = driver.find_element_by_class_name('db-exclude-keyword-search')
        element.send_keys(Keys.CONTROL + "a")
        element.send_keys(Keys.DELETE)
        excludedKeyword = str(searchType['Excluded Keyword'])
        element.send_keys(excludedKeyword)


class MinMaxInputFieldDriver:
    def enterAllData(self, searchType):
        self.enterDATA(searchType)

    def enterDATA(self, searchType):
        allElem = driver.find_elements_by_xpath(".//span[@class = 'db-filter-content']")
        for item in allElem:
            for key in searchType.keys():
                print("-------" + str(searchType.get(key)[0]))
                if type(searchType.get(key)[0]) is int:
                    if str(key) in item.text:
                        if searchType[key] is not None and searchType.get(key)[0] != 0:
                            self.enterMin(searchType.get(key)[0], item, MINMAXINDEXMAP[key])
                        if len(searchType[key]) >= 2 and searchType.get(key)[1] != 0:
                            self.enterMax(searchType.get(key)[1], item, MINMAXINDEXMAP[key])

    @staticmethod
    def enterMin(valueToEnter, item, index):
        inputMinBox = item.find_elements_by_xpath("//input[@placeholder = 'min']")
        inputMinBox[index].send_keys(Keys.CONTROL + "a")
        inputMinBox[index].send_keys(Keys.DELETE)
        inputMinBox[index].send_keys(str(valueToEnter))

    @staticmethod
    def enterMax(valueToEnter, item, index):
        inputMaxBox = item.find_elements_by_xpath("//input[@placeholder = 'max']")
        inputMaxBox[index].send_keys(Keys.CONTROL + "a")
        inputMaxBox[index].send_keys(Keys.DELETE)
        inputMaxBox[index].send_keys(str(valueToEnter))

if __name__ == '__main__':
    AutomateJSWork()

