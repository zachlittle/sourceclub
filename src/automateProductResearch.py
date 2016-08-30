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
    SEARCHKEYWORD = ['dogs']
    EXCLUDEKEYWORD = ['cats, goats, mothar']
    PRICERANGE = ['10', '15']
    NET = ['8', '']
    RANK = ['1', '10000']
    ESTIMATEDSALES = ['200', '4000']
    ESTIMATEDREVENUE = ['3000', '200']
    REVIEWS = ['300', '500']
    RATING = ['50', '75']
    WEIGHT = ['5', '10']
    NUMBEROFSELLERS = ['10', '100']
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
    EXCLUDEKEYWORD = ['cats, goats, mothar']
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
        AutomateJSWork.Login()
        AutomateJSWork.InputSearchTypeData(LOWPRICEPRODUCT)
        AutomateJSWork.Logout()
        driver.quit()
        '''        except:
            e = sys.exc_info()[0]
            print('Exception: ' + str(e))'''

    @staticmethod
    def Login():
        driver.get('https://www.junglescout.com/')
        driver.implicitly_wait(100)
        driver.find_element_by_id('menu-item-4544').click()
        inputElementUsername = driver.find_element_by_id('user_login')
        inputElementUsername.send_keys(LOGINCREDENTIALS['Username'])
        inputElementPassword = driver.find_element_by_id('user_password')
        inputElementPassword.send_keys(LOGINCREDENTIALS['Password'])
        driver.find_element_by_name('commit').click()

    @staticmethod
    def Logout():
        pass
        # add find dropdown and select logout

    def clickSearch(self):
        element = driver.find_element_by_class_name('make-search')
        element.click()

    def InputSearchTypeData(searchType):
        print("CURRENTLY RUNNING " + str(searchType) + " SEARCH")
        driver.implicitly_wait(10)
        driver.find_element_by_class_name('database-nav').click()
        AutomateJSWork.inputData(searchType)

    def inputData(searchType):
        AutomateJSWork.clickDesiredCategories(searchType)
        MinMaxInputFieldDriver.enterAllData(searchType)

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

    def enterPriceRange(searchType):
        allElem = driver.find_elements_by_xpath(".//span[@class = 'db-filter-content']")
        for item in allElem:
            if "Price" in item.text:
                if searchType.PRICERANGE[0] is not None:
                    MinMaxInputFieldDriver.enterMin(item, searchType, MINMAXINDEXMAP['PRICERANGE'])
                if len(searchType.PRICERANGE) >= 2:
                    MinMaxInputFieldDriver.enterMax(item, searchType, MINMAXINDEXMAP['PRICERANGE'])

    def enterNetRange(searchType):
        allElem = driver.find_elements_by_xpath(".//span[@class = 'db-filter-content']")
        for item in allElem:
            print(item.text)
            if "Net" in item.text:
                if searchType.NET[0] is not None:
                    MinMaxInputFieldDriver.enterMin(item, searchType, MINMAXINDEXMAP['NET'])
                if len(searchType.NET) >= 2:
                    MinMaxInputFieldDriver.enterMax(item, searchType, MINMAXINDEXMAP['NET'])

    def enterRankRange(searchType):
        allElem = driver.find_elements_by_xpath(".//span[@class = 'db-filter-content']")
        for item in allElem:
            print(item.text)
            if "Rank" in item.text:
                if searchType.RANK[0] is not None:
                    MinMaxInputFieldDriver.enterMin(item, searchType, MINMAXINDEXMAP['RANK'])
                if len(searchType.RANK) >= 2:
                    MinMaxInputFieldDriver.enterMax(item, searchType, MINMAXINDEXMAP['RANK'])

    def enterEstimatedSalesRange(searchType):
        allElem = driver.find_elements_by_xpath(".//span[@class = 'db-filter-content']")
        for item in allElem:
            print(item.text)
            if "Est. Sales" in item.text:
                if searchType.ESTIMATEDSALES[0] is not None:
                    MinMaxInputFieldDriver.enterMin(item, searchType, MINMAXINDEXMAP['ESTIMATEDSALES'])
                if len(searchType.ESTIMATEDSALES) >= 2:
                    MinMaxInputFieldDriver.enterMax(item, searchType, MINMAXINDEXMAP['ESTIMATEDSALES'])

    def enterEstimatedRevenueRange(searchType):
        allElem = driver.find_elements_by_xpath(".//span[@class = 'db-filter-content']")
        for item in allElem:
            print(item.text)
            if "Est. Revenue" in item.text:
                if searchType.ESTIMATEDREVENUE[0] is not None:
                    MinMaxInputFieldDriver.enterMin(item, searchType, MINMAXINDEXMAP['ESTIMATEDREVENUE'])
                if len(searchType.ESTIMATEDREVENUE) >= 2:
                    MinMaxInputFieldDriver.enterMax(item, searchType, MINMAXINDEXMAP['ESTIMATEDREVENUE'])

    def enterReviewsRange(searchType):
        allElem = driver.find_elements_by_xpath(".//span[@class = 'db-filter-content']")
        for item in allElem:
            print(item.text)
            if "Reviews" in item.text:
                if searchType.REVIEWS[0] is not None:
                    MinMaxInputFieldDriver.enterMin(item, searchType, MINMAXINDEXMAP['REVIEWS'])
                if len(searchType.REVIEWS) >= 2:
                    MinMaxInputFieldDriver.enterMax(item, searchType, MINMAXINDEXMAP['REVIEWS'])

    def enterRatingRange(searchType):
        allElem = driver.find_elements_by_xpath(".//span[@class = 'db-filter-content']")
        for item in allElem:
            print(item.text)
            if "Rating" in item.text:
                if searchType.RATING[0] is not None:
                    MinMaxInputFieldDriver.enterMin(item, searchType, MINMAXINDEXMAP['RATING'])
                if len(searchType.RATING) >= 2:
                    MinMaxInputFieldDriver.enterMax(item, searchType, MINMAXINDEXMAP['RATING'])

    def enterMin(self, searchType, index):
        inputMinBox = self.find_elements_by_xpath("//input[@placeholder = 'min']")
        inputMinBox[index].send_keys(Keys.CONTROL + "a")
        inputMinBox[index].send_keys(Keys.DELETE)
        inputValMin = str(searchType.NET[0])
        inputMinBox[index].send_keys(inputValMin)

    def enterMax(self, searchType, index):
        inputMaxBox = self.find_elements_by_xpath("//input[@placeholder = 'max']")
        inputMaxBox[index].send_keys(Keys.CONTROL + "a")
        inputMaxBox[index].send_keys(Keys.DELETE)
        inputValMax = str(searchType.NET[1])
        inputMaxBox[index].send_keys(inputValMax)


if __name__ == '__main__':
    AutomateJSWork()
