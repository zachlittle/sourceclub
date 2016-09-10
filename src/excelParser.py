from openpyxl import load_workbook
import glob, os, shutil
import ast
import xlwt
import xlrd
import datetime

global mostRecentExcelFile
global currentSheet
productData = {}

CONSTRAINTS = {'Price': {'Min': 20, 'Max': 100},
               'Fees': {},
               'Net': {'Min': 10},
               'Weight': {'Max': 10},
               'Reviews': {'Max': 200},
               'AverageRating': {},
               'Rank': {},
               'EstimatedMonthlySales': {'Min': 300},
               'EstimatedMonthlyRevenue': {},
               'NumberOfSellers': {'Max': 15}
              }

class excelHandler:

    def __init__(self):
        self.findAndMoveExcelFiles()
        self.renameCVSFilesToXLSX()
        self.pickMostRecentExcelFileToParse()
        self.selectSheetByFilePath()
        #self.packageAllFilesRelatedIntoOneWorkbook()

    @staticmethod
    def selectSheetByFilePath():
        global currentSheet
        excelFileName = 'C:\\Git\\sourceclub\\src\\Excel Files\\' + str(mostRecentExcelFile)
        wb2 = load_workbook(excelFileName)
        sheetName = wb2.get_sheet_names()
        currentSheet = wb2.get_sheet_by_name(sheetName[0])
        sheetToListGenerator(currentSheet)

    @staticmethod
    def pickMostRecentExcelFileToParse():
        import os
        dir = 'C:\\Git\\sourceclub\\src\\Excel Files\\'
        guyList = []
        maximumModTime = 0
        for file in os.listdir(dir):
            if file.endswith(".xlsx"):
                guyList.append(os.path.getmtime(str(dir) + "" + str(file)))

        for modTime in guyList:
            if modTime > maximumModTime:
                maximumModTime = modTime
        import datetime
        dt_obj = datetime.datetime.utcfromtimestamp(maximumModTime)
        print("Most Recent Excel File To Parse: " + str(dt_obj))
        excelHandler.helperMethodCuzImToBakedToThinkOfAName(dt_obj, dir)

    @staticmethod
    def helperMethodCuzImToBakedToThinkOfAName(dateTime, dir):
        global mostRecentExcelFile
        for file in os.listdir(dir):
            print("----------------------------------------------------------------------------------")
            currentFileDateTime = datetime.datetime.utcfromtimestamp(os.path.getmtime(str(dir) + "" + str(file)))
            if currentFileDateTime == dateTime:
                print("Most Recent File: " + file)
                mostRecentExcelFile = file

    @staticmethod
    def findAndMoveExcelFiles():
        source_dir = 'C:\\Users\maxsm_000\Downloads'
        dest_dir = 'C:\Git\sourceclub\src\Excel Files'
        for file in os.listdir(source_dir):
            if file.endswith(".csv"):
                shutil.copy2(source_dir + '\\' + file, dest_dir)
                os.remove(source_dir + '\\' + file)

    @staticmethod
    def renameCVSFilesToXLSX():
        source_dir = 'C:\\Git\\sourceclub\\src\\Excel Files'
        for file in os.listdir(source_dir):
            if file.endswith(".csv"):
                pre, ext = os.path.splitext(file)
                os.rename(source_dir + '\\' + file, 'Excel Files\\' + pre + '.xlsx')

    # def packageAllFilesRelatedIntoOneWorkbook(self):
    #     workbooksToCombine = []
    #     wkbk = xlwt.Workbook()
    #     outsheet = wkbk.add_sheet('Sheet1')
    #     outrow_idx = 0
    #     source_dir = 'C:\\Git\\sourceclub\\src\\Excel Files'
    #     for file in os.listdir(source_dir):
    #         if 'Jungle' in file:
    #             workbooksToCombine.append(file)
    #     for workbook in workbooksToCombine:
    #         print(source_dir + '\\' + workbook)
    #         wb2 = load_workbook(source_dir + '\\' + workbook)
    #         insheet = wb2.worksheets[0]
    #         for row_idx in insheet['A1':'B1']:
    #             for col_idx in range(insheet):
    #                 outsheet.write(outrow_idx, col_idx,
    #                                insheet.cell_value(row_idx, col_idx))
    #             outrow_idx += 1
    #     date = datetime.date
    #     wkbk.save(source_dir + '\\combinedFiles' + date + '.xlsx')

class sheetToListGenerator:
    primaryKeyColumn = 'A'
    primaryKeyList = {}
    filteredProducts = [];

    def __init__(self, sheet):
        self.createPrimaryKeyBasedOnColumn(sheet)
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
        for keyIndex, key in enumerate(self.primaryKeyList):
            for category in productData[key]:
                if CONSTRAINTS.get(str(category)) is not None:
                    if type(productData[key][str(category)]) is not str:
                        if 'Min' in CONSTRAINTS[str(category)].keys():
                            if productData[key][str(category)] < CONSTRAINTS[str(category)]['Min']:
                                productsToRemove.append(key)
                                break
                        if 'Max' in CONSTRAINTS[str(category)].keys():
                            if productData[key][str(category)] > CONSTRAINTS[str(category)]['Max']:
                                productsToRemove.append(key)
                                break
        self.removeProductsByIndex(productsToRemove)

    def removeProductsByIndex(self, productsToRemove):
        for product in productsToRemove:
            self.filteredProducts.append(product)
            del productData[product]
            del self.primaryKeyList[product]

    def createPrimaryKeyBasedOnColumn(self, sheet):
        for i, row in enumerate(range(4, sheet.max_row + 1)):
            primaryKey = currentSheet['' + self.primaryKeyColumn + '' + str(row)].value
            productData.setdefault('' + primaryKey + '', {})
            self.primaryKeyList.setdefault('' + primaryKey + '', {})

    def populateDataStructure(self):
        self.createDataItem('Name', 'B')
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
            print('ASIN: ' + key)
            print('Name: ' + productData[key]['Name'])
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

excelHandler()