from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import sys


path = 'C:/Git/sourceclub/src/chromedriver.exe'
driver = webdriver.Chrome(path)

LOGINCREDENTIALS = {'Username': 'sourceclubcommerce@gmail.com',
                    'Password': 'chronicboys760'}

MINMAXINDEXMAP = {'PRICERANGE': 0,
                  'NET': 1,
                  'RANK': 2,
                  'ESTIMATEDSALES': 3,
                  'ESTIMATEDREVENUE': 4,
                  'REVIEWS': 5,
                  'RATING': 6,
                  'WEIGHT': 7,
                  'NUMBEROFSELLERS': 8,
                  'LISTINGQUALITY': 9}


class LOWPRICEPRODUCT:
    CATEGORIES = ['Home Improvement', 'Automotive', 'Clothing', 'Office Products']
    PRODUCTTIER = ['Standard', 'Oversize']
    SELLER = ['Amazon', 'Fulfilled By Amazon', 'Fulfilled By Merchant']
    SEARCHKEYWORD = ['']
    EXCLUDEKEYWORD = ['bible']
    PRICERANGE = ['20', '100']
    NET = ['10']
    RANK = ['', '']
    ESTIMATEDSALES = ['300']
    ESTIMATEDREVENUE = ['', '']
    REVIEWS = ['', '200']
    RATING = ['', '']
    WEIGHT = ['','10']
    NUMBEROFSELLERS = ['', '15']
    LISTINGQUALITY = ['5', '90']


class MEDIUMPRICEPRODUCT:
    CATEGORIES = ['Home Improvement', 'Automotive', 'Clothing', 'Office Products']
    PRODUCTTIER = ['Standard', 'Oversize']
    SELLER = ['Amazon', 'Fulfilled By Amazon', 'Fulfilled By Merchant']
    SEARCHKEYWORD = ['dogs']
    EXCLUDEKEYWORD = ['cats, goats, mothar']
    PRICERANGE = ['21', '50']
    NET = ['8']
    RANK = ['1', '2']
    ESTIMATEDSALES = ['300', '500']
    ESTIMATEDREVENUE = ['100', '200']
    REVIEWS = ['300', '500']
    RATING = ['50', '75']
    WEIGHT = ['5', '10']
    NUMBEROFSELLERS = ['10', '100']
    LISTINGQUALITY = ['5', '90']


class HIGHPRICEPRODUCT:
    CATEGORIES = ['Home Improvement', 'Automotive', 'Clothing', 'Office Products']
    PRODUCTTIER = ['Standard', 'Oversize']
    SELLER = ['Amazon', 'Fulfilled By Amazon', 'Fulfilled By Merchant']
    SEARCHKEYWORD = ['dogs']
    EXCLUDEKEYWORD = ['nike']
    PRICERANGE = ['51', '80']
    NET = ['8']
    RANK = ['1', '2']
    ESTIMATEDSALES = ['300', '500']
    ESTIMATEDREVENUE = ['100', '200']
    REVIEWS = ['300', '500']
    RATING = ['50', '75']
    WEIGHT = ['5', '10']
    NUMBEROFSELLERS = ['10', '100']
    LISTINGQUALITY = ['5', '90']


class AutomateJSWork:
    def __init__(self):
        #try:
        self.Login()
        self.clickDatabase()
        self.InputSearchTypeData(LOWPRICEPRODUCT)
        self.Logout()
        driver.quit()
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
    def clickSearch(self):
        element = driver.find_element_by_class_name('make-search')
        element.click()


    def InputSearchTypeData(self, searchType):
        print("CURRENTLY RUNNING " + str(searchType) + " SEARCH")
        driver.implicitly_wait(10)
        driver.find_element_by_class_name('filter-reset-btn').click()
        AutomateJSWork.inputData(searchType)
        AutomateJSWork.clickResultsPerPage()
        AutomateJSWork.clickSearch()
        AutomateJSWork.clickExportToExcel()

    @staticmethod
    def clickDatabase():
        driver.find_element_by_class_name('database-nav').click()

    @staticmethod
    def clickResultsPerPage():
        element = driver.find_element_by_class_name('change-page-counter')
        element.click()
        element2 = element.find_element_by_xpath("//option[@value = '200']")
        element2.click()

    def inputData(searchType):
        AutomateJSWork.clickDesiredCategories(searchType)
        MinMaxInputFieldDriver.enterAllData(searchType)
        AutomateJSWork.clickProductTier(searchType)
        AutomateJSWork.clickSeller(searchType)
        AutomateJSWork.enterSearchKeyword(searchType)
        AutomateJSWork.enterExcludedKeyword(searchType)

    def clickDesiredCategories(searchType):
        necItemsToClick = len(searchType.CATEGORIES)
        for listItem in searchType.CATEGORIES:
            if necItemsToClick > 0:
                allElem = driver.find_elements_by_xpath(".//span[@class = 'category-label']")
                for item in allElem:
                    if listItem in item.text:
                        item.click()
                        necItemsToClick -= 1
            else:
                break

    def clickProductTier(searchType):
        necItemsToClick = len(searchType.PRODUCTTIER)
        for listItem in searchType.PRODUCTTIER:
            if necItemsToClick > 0:
                allElem = driver.find_elements_by_xpath(".//div[@class = 'category-label-content']")
                for item in allElem:
                    if listItem in item.text:
                        item.click()
                        necItemsToClick -= 1
            else:
                break

    def clickSeller(searchType):
        necItemsToClick = len(searchType.SELLER)
        for listItem in searchType.SELLER:
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

    def enterSearchKeyword(searchType):
        element = driver.find_element_by_class_name('db-keyword-search')
        element.send_keys(Keys.CONTROL + "a")
        element.send_keys(Keys.DELETE)
        searchKeyword = str(searchType.SEARCHKEYWORD[0])
        element.send_keys(searchKeyword)

    def enterExcludedKeyword(searchType):
        element = driver.find_element_by_class_name('db-exclude-keyword-search')
        element.send_keys(Keys.CONTROL + "a")
        element.send_keys(Keys.DELETE)
        excludedKeyword = str(searchType.EXCLUDEKEYWORD[0])
        element.send_keys(excludedKeyword)


class MinMaxInputFieldDriver:
    def enterAllData(searchType):
        MinMaxInputFieldDriver.enterPriceRange(searchType)
        MinMaxInputFieldDriver.enterEstimatedRevenueRange(searchType)
        MinMaxInputFieldDriver.enterNetRange(searchType)
        MinMaxInputFieldDriver.enterEstimatedSalesRange(searchType)
        MinMaxInputFieldDriver.enterRankRange(searchType)
        MinMaxInputFieldDriver.enterEstimatedRevenueRange(searchType)
        MinMaxInputFieldDriver.enterReviewsRange(searchType)
        MinMaxInputFieldDriver.enterRatingRange(searchType)
        MinMaxInputFieldDriver.enterWeightRange(searchType)
        MinMaxInputFieldDriver.enterNumberOfSellersRange(searchType)
        MinMaxInputFieldDriver.enterListingQualityRange(searchType)

    def enterPriceRange(searchType):
        allElem = driver.find_elements_by_xpath(".//span[@class = 'db-filter-content']")
        for item in allElem:
            if "Price" in item.text:
                if searchType.PRICERANGE[0] is not None:
                    MinMaxInputFieldDriver.enterMin(searchType.PRICERANGE[0], item, MINMAXINDEXMAP['PRICERANGE'])
                if len(searchType.PRICERANGE) >= 2:
                    MinMaxInputFieldDriver.enterMax(searchType.PRICERANGE[1], item, MINMAXINDEXMAP['PRICERANGE'])

    def enterNetRange(searchType):
        allElem = driver.find_elements_by_xpath(".//span[@class = 'db-filter-content']")
        for item in allElem:
            if "Net" in item.text:
                if searchType.NET[0] is not None:
                    MinMaxInputFieldDriver.enterMin(searchType.NET[0], item, MINMAXINDEXMAP['NET'])
                if len(searchType.NET) >= 2:
                    MinMaxInputFieldDriver.enterMax(searchType.NET[1], item, MINMAXINDEXMAP['NET'])

    def enterRankRange(searchType):
        allElem = driver.find_elements_by_xpath(".//span[@class = 'db-filter-content']")
        for item in allElem:
            if "Rank" in item.text:
                if searchType.RANK[0] is not None:
                    MinMaxInputFieldDriver.enterMin(searchType.RANK[0], item, MINMAXINDEXMAP['RANK'])
                if len(searchType.RANK) >= 2:
                    MinMaxInputFieldDriver.enterMax(searchType.RANK[1], item, MINMAXINDEXMAP['RANK'])

    def enterEstimatedSalesRange(searchType):
        allElem = driver.find_elements_by_xpath(".//span[@class = 'db-filter-content']")
        for item in allElem:
            if "Est. Sales" in item.text:
                if searchType.ESTIMATEDSALES[0] is not None:
                    MinMaxInputFieldDriver.enterMin(searchType.ESTIMATEDSALES[0], item, MINMAXINDEXMAP['ESTIMATEDSALES'])
                if len(searchType.ESTIMATEDSALES) >= 2:
                    MinMaxInputFieldDriver.enterMax(searchType.ESTIMATEDSALES[1], item, MINMAXINDEXMAP['ESTIMATEDSALES'])

    def enterEstimatedRevenueRange(searchType):
        allElem = driver.find_elements_by_xpath(".//span[@class = 'db-filter-content']")
        for item in allElem:
            if "Est. Rev" in item.text:
                if searchType.ESTIMATEDREVENUE[0] is not None:
                    MinMaxInputFieldDriver.enterMin(searchType.ESTIMATEDREVENUE[0], item, MINMAXINDEXMAP['ESTIMATEDREVENUE'])
                if len(searchType.ESTIMATEDREVENUE) >= 2:
                    MinMaxInputFieldDriver.enterMax(searchType.ESTIMATEDREVENUE[1], item, MINMAXINDEXMAP['ESTIMATEDREVENUE'])

    def enterReviewsRange(searchType):
        allElem = driver.find_elements_by_xpath(".//span[@class = 'db-filter-content']")
        for item in allElem:
            if "Reviews" in item.text:
                if searchType.REVIEWS[0] is not None:
                    MinMaxInputFieldDriver.enterMin(searchType.REVIEWS[0], item, MINMAXINDEXMAP['REVIEWS'])
                if len(searchType.REVIEWS) >= 2:
                    MinMaxInputFieldDriver.enterMax(searchType.REVIEWS[0], item, MINMAXINDEXMAP['REVIEWS'])

    def enterRatingRange(searchType):
        allElem = driver.find_elements_by_xpath(".//span[@class = 'db-filter-content']")
        for item in allElem:
            if "Rating" in item.text:
                if searchType.RATING[0] is not None:
                    MinMaxInputFieldDriver.enterMin(searchType.RATING[0], item, MINMAXINDEXMAP['RATING'])
                if len(searchType.RATING) >= 2:
                    MinMaxInputFieldDriver.enterMax(searchType.RATING[1], item, MINMAXINDEXMAP['RATING'])

    def enterWeightRange(searchType):
        allElem = driver.find_elements_by_xpath(".//span[@class = 'db-filter-content']")
        for item in allElem:
            if "Weight" in item.text:
                if searchType.WEIGHT[0] is not None:
                    MinMaxInputFieldDriver.enterMin(searchType.WEIGHT[0], item, MINMAXINDEXMAP['WEIGHT'])
                if len(searchType.WEIGHT) >= 2:
                    MinMaxInputFieldDriver.enterMax(searchType.WEIGHT[1], item, MINMAXINDEXMAP['WEIGHT'])

    def enterNumberOfSellersRange(searchType):
        allElem = driver.find_elements_by_xpath(".//span[@class = 'db-filter-content']")
        for item in allElem:
            if "No. Sellers" in item.text:
                if searchType.NUMBEROFSELLERS[0] is not None:
                    MinMaxInputFieldDriver.enterMin(searchType.NUMBEROFSELLERS[0], item, MINMAXINDEXMAP['NUMBEROFSELLERS'])
                if len(searchType.NUMBEROFSELLERS) >= 2:
                    MinMaxInputFieldDriver.enterMax(searchType.NUMBEROFSELLERS[1], item, MINMAXINDEXMAP['NUMBEROFSELLERS'])

    def enterListingQualityRange(searchType):
        allElem = driver.find_elements_by_xpath(".//span[@class = 'db-filter-content']")
        for item in allElem:
            if "Listing Quality" in item.text:
                if searchType.LISTINGQUALITY[0] is not None:
                    MinMaxInputFieldDriver.enterMin(searchType.LISTINGQUALITY[0], item, MINMAXINDEXMAP['LISTINGQUALITY'])
                if len(searchType.LISTINGQUALITY) >= 2:
                    MinMaxInputFieldDriver.enterMax(searchType.LISTINGQUALITY[1], item, MINMAXINDEXMAP['LISTINGQUALITY'])

    def enterMin(searchType, item, index):
        inputMinBox = item.find_elements_by_xpath("//input[@placeholder = 'min']")
        inputMinBox[index].send_keys(Keys.CONTROL + "a")
        inputMinBox[index].send_keys(Keys.DELETE)
        inputValMin = str(searchType)
        inputMinBox[index].send_keys(inputValMin)

    def enterMax(searchType, item, index):
        inputMaxBox = item.find_elements_by_xpath("//input[@placeholder = 'max']")
        inputMaxBox[index].send_keys(Keys.CONTROL + "a")
        inputMaxBox[index].send_keys(Keys.DELETE)
        inputValMax = str(searchType)
        inputMaxBox[index].send_keys(inputValMax)


if __name__ == '__main__':
    AutomateJSWork()

