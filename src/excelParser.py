from openpyxl import load_workbook
import glob, os, shutil
import ast

excelFileName = 'Excel Files\productSearch9_3Page1.xlsx'
wb2 = load_workbook(excelFileName)
currentSheet = wb2.get_sheet_by_name('Sheet1')
productData = {}

CONSTRAINTS = {'Price': {'Min': 1, 'Max': 1000},
               'Fees': {'Min': 1, 'Max': 10},
               'Net': {'Min': 1, 'Max': 1000},
               'Weight': {'Min': .1, 'Max': 10},
               'Reviews': {'Min': 1, 'Max': 500},
               'AverageRating': {'Min': 1, 'Max': 6},
               'Rank': {'Min': 1, 'Max': 1000000000},
               'EstimatedMonthlySales': {'Min': 1, 'Max': 100000},
               'EstimatedMonthlyRevenue': {'Min': 1, 'Max': 1000000},
               'NumberOfSellers': {'Min': 1, 'Max': 1000}
              }

class sheetToListGenerator():
    primaryKeyColumn = 'B'
    primaryKeyList = {}
    filteredProducts = [];

    def __init__(self):
        self.findAndMoveExcelFiles()
        self.createPrimaryKeyBasedOnColumn()
        print(len(self.primaryKeyList))
        self.populateDataStructure()
        self.outputStructure()
        self.filterByConstraints()
        self.printStats()

    def printStats(self):
        print('--------------------FILTERSTATS---------------------')
        print('# of products filtered due to failure to meet constraints: ' + str(len(self.filteredProducts)))
        print("Name's of products filtered: ")
        for product in self.filteredProducts:
            print(product)
        print('# of viable products remaining: ' + str(len(productData)))
        print("Name's of products remaining: ")
        for product in productData:
            print(product)
        print('--------------------FILTERSTATS---------------------')


    def filterByConstraints(self):
        productsToRemove = []
        productsToRemoveByIndex = []
        for keyIndex, key in enumerate(self.primaryKeyList):
            for category in productData[key]:
                if CONSTRAINTS.get(str(category)) is not None:
                    if type(productData[key][str(category)]) is not str:
                        if productData[key][str(category)] < CONSTRAINTS[str(category)]['Min']:
                            productsToRemove.append(key)
                            productsToRemoveByIndex.append(keyIndex)
                            break
                        elif productData[key][str(category)] > CONSTRAINTS[str(category)]['Max']:
                            productsToRemove.append(key)
                            productsToRemoveByIndex.append(keyIndex)
                            break
        self.removeProductsByIndex(productsToRemove, productsToRemoveByIndex)

    def removeProductsByIndex(self, productsToRemove, productsToRemoveByIndex):
        for product in productsToRemove:
            self.filteredProducts.append(product)
            del productData[product]
            del self.primaryKeyList[product]



    @staticmethod
    def findAndMoveExcelFiles():
        source_dir = 'C:\\Users\maxsm_000\Downloads'
        dest_dir = 'C:\Git\sourceclub\src\Excel Files'
        for file in os.listdir(source_dir):
            if file.endswith(".csv"):
                shutil.copy2(source_dir + '\\' + file, dest_dir)
                os.remove(source_dir + '\\' + file)

    def createPrimaryKeyBasedOnColumn(self):
        for i, row in enumerate(range(4, currentSheet.max_row + 1)):
            primaryKey = currentSheet['' + self.primaryKeyColumn + '' + str(row)].value
            productData.setdefault('' + primaryKey + '', {})
            self.primaryKeyList.setdefault('' + primaryKey + '', {})

    def populateDataStructure(self):
        self.createDataItem('ASIN', 'A')
        self.createDataItem('Price', 'F')
        self.createDataItem('Brand', 'C')
        self.createDataItem('Seller', 'D')
        self.createDataItem('Category', 'E')
        self.createDataItem('Fees', 'G')
        self.createDataItem('Net', 'H')
        self.createDataItem('Weight', 'I')
        self.createDataItem('ProductTier', 'J')
        self.createDataItem('Reviews', 'K')
        self.createDataItem('AverageRating', 'L')
        self.createDataItem('Rank', 'M')
        self.createDataItem('EstimatedMonthlySales', 'N')
        self.createDataItem('EstimatedMonthlyRevenue', 'O')
        self.createDataItem('NumberOfSellers', 'Q')

    @staticmethod
    def createDataItem(itemToAdd, columnValue):
        for row in range(4, currentSheet.max_row + 1):
            primaryKey = currentSheet['' + sheetToListGenerator.primaryKeyColumn + '' + str(row)].value
            itemValues = currentSheet['' + columnValue + '' + str(row)].value
            productData[primaryKey].setdefault(str(itemToAdd), itemValues)

    @staticmethod
    def outputStructure():
        for i, key in enumerate(sheetToListGenerator.primaryKeyList):
            print('--------------Item' + str(i) + '------------------')
            print('Name: ' + key)
            print('ASIN: ' + productData[key]['ASIN'])
            print('Brand: ' + str(productData[key]['Brand']))
            print('Price: ' + str(productData[key]['Price']))
            print('Seller: ' + str(productData[key]['Seller']))
            print('Category: ' + str(productData[key]['Category']))
            print('Fees: ' + str(productData[key]['Fees']))
            print('Net: ' + str(productData[key]['Net']))
            print('Weight: ' + str(productData[key]['Weight']))
            print('ProductTier: ' + str(productData[key]['ProductTier']))
            print('Reviews: ' + str(productData[key]['Reviews']))
            print('AverageRating: ' + str(productData[key]['AverageRating']))
            print('Rank: ' + str(productData[key]['Rank']))
            print('EstimatedMonthlySales: ' + str(productData[key]['EstimatedMonthlySales']))
            print('EstimatedMonthlyRevenue: ' + str(productData[key]['EstimatedMonthlyRevenue']))
            print('NumberOfSellers: ' + str(productData[key]['NumberOfSellers']))

sheetToListGenerator()