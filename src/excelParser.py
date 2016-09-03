from openpyxl import load_workbook
import glob, os, shutil
excelFileName = 'Excel Files\productSearch9_3Page1.xlsx'
wb2 = load_workbook(excelFileName)
currentSheet = wb2.get_sheet_by_name('Sheet1')
productData = {}

class sheetToListGenerator():
    primaryKeyColumn = 'B'
    primaryKeyList = []

    def __init__(self):
        self.findAndMoveExcelFiles()
        self.createPrimaryKey(sheetToListGenerator.primaryKeyColumn)
        self.populateDataStructure()
        self.outputStructure()

    @staticmethod
    def findAndMoveExcelFiles():
        source_dir = 'C:\\Users\maxsm_000\Downloads'
        dest_dir = 'C:\Git\sourceclub\src\Excel Files'
        for file in os.listdir(source_dir):
            if file.endswith(".csv"):
                shutil.copy2(source_dir + '\\' + file, dest_dir)
                os.remove(source_dir + '\\' + file)

    @staticmethod
    def createPrimaryKey(primaryKeyColumnValue):
        for i, row in enumerate(range(4, currentSheet.max_row + 1)):
            primaryKey = currentSheet['' + primaryKeyColumnValue + '' + str(row)].value
            productData.setdefault('' + primaryKey + '', {})
            sheetToListGenerator.primaryKeyList.append(primaryKey)

    @staticmethod
    def populateDataStructure():
        sheetToListGenerator.createDataItem('ASIN', 'A')
        sheetToListGenerator.createDataItem('Price', 'F')
        sheetToListGenerator.createDataItem('Brand', 'C')
        sheetToListGenerator.createDataItem('Seller', 'D')
        sheetToListGenerator.createDataItem('Category', 'E')
        sheetToListGenerator.createDataItem('Fees', 'G')
        sheetToListGenerator.createDataItem('Net', 'H')
        sheetToListGenerator.createDataItem('Weight', 'I')
        sheetToListGenerator.createDataItem('ProductTier', 'J')
        sheetToListGenerator.createDataItem('Reviews', 'K')
        sheetToListGenerator.createDataItem('AverageRating', 'L')
        sheetToListGenerator.createDataItem('Rank', 'M')
        sheetToListGenerator.createDataItem('EstimatedMonthlySales', 'N')
        sheetToListGenerator.createDataItem('EstimatedMonthlyRevenue', 'O')
        sheetToListGenerator.createDataItem('NumberOfSellers', 'Q')

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