import xlrd

# ***************************
# managing Excel workbooks *
# ***************************

class excelWB:
    def __init__(self, pathToExcelFile, activeSheetName=''):
        self.workbook = self.openWorkBook(pathToExcelFile)
        self.sheetnamesList = self.getSheetNames()
        if activeSheetName == '':
            self.activeSheet = self.getSheet(self.sheetnamesList[0])  # we initialize on first sheet
        else:
            self.activeSheet = self.getSheet(activeSheetName)

    def getSheetNames(self):
        # returns a list of sheet names in workbook
        sheetList = []
        for s in self.workbook.sheets():
            sheetList.append(s.name)
        return sheetList

    def getSheet(self, sheetName):
        # returns a sheet object sheetname in workbook
        return self.workbook.sheet_by_name(sheetName)

    def openWorkBook(self, path):
        # Open and read an Excel file
        book = xlrd.open_workbook(path)
        return book

    def getCellNumberByStringValue(self, key):  # ???
        # returns the first cell with value == key
        cellNumber = [-1, -1]  # returned when value is not found

        for row in range(self.activeSheet.nrows):
            for col in range(self.activeSheet.ncols):
                if self.activeSheet.cell_value(row, col) == key:
                    cellNumber = [row, col]
        return cellNumber

    def getMultipleCellNumbersByStringValue(self, key):
        # keys are not necessarily unique, so we return a list
        cellNumbers = []
        for row in range(self.activeSheet.nrows):
            for col in range(self.activeSheet.ncols):
                if self.activeSheet.cell_value(row, col) == key:
                    cellNumbers.append([row, col])
        return cellNumbers

    def getFieldValue(self, keyName, templateSheet):
        # This routine is only to be used to get the general lab info;
        # i.e, equipment information is differently structured in the Excel-template
        # pre-assumptions:
        # - field names are unique in the template
        # - value is to be found in the cell next to the right of the label-field cell
        # To remove default instruction values:
        # - value is checked against value from template:
        # - if equal empty string is returned

        cellNumber = self.getCellNumberByStringValue(keyName)
        if len(cellNumber) != 0:
            returnVal = self.activeSheet.cell(int(cellNumber[0]), int(cellNumber[1] + 1)).value
        else:
            returnVal = ''
        # check against default value from template:
        templateCellNumber = self.getCellNumberByStringValue(keyName)
        if len(templateCellNumber) != 0:
            templateVal = templateSheet.cell(int(templateCellNumber[0]), int(templateCellNumber[1] + 1)).value
        else:
            templateVal = ''
        if returnVal == templateVal:
            returnVal = ''
        return returnVal

    def getFieldValueByAdjacentCellNum(self, cellNum):
        # pre-assumptions:
        # value is to be found in the cell next to the right
        returnVal = self.activeSheet.cell(int(cellNum[0]), int(cellNum[1] + 1)).value
        return returnVal

    def getFieldValueByCellNum(self, cellNum, templateDefaultValuesList):
        try:
            returnVal = self.activeSheet.cell(int(cellNum[0]), int(cellNum[1])).value
            # skip default values from template
            if returnVal in templateDefaultValuesList:
                returnVal = ''
        except:
            returnVal = ''

        return returnVal

    def getListOfMergedCellValues(self):
        # (horizontally) merged cells are separator rows in the template
        # returns a list of (values for) these rows
        mergedCellsList = self.activeSheet.merged_cells
        mergedValues = []
        for c in mergedCellsList:
            mergedValues.append(self.activeSheet.cell(c[0], 0).value)
        return mergedValues

    def checkOnEmptyRow(self, nrow):
        empty = True
        row = self.activeSheet.row(nrow)
        for cell in row:
            if not (cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK)):
                empty = False
        return empty

