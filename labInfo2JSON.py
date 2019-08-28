#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Aug 15 12:34:05 2019

@author: Otto Lange and the EPOS Multi-scale Laboratories team

Version 0.1

This inital version is meant to handle 4 different Excel-templates, i.e. one for each MSL subdomain.
The initial templates were partially inconsistently structured. If we proceed by using Excel-templates
we may move to a single highly structured template. The 'single-usage' code will then be translated
into OO-code ready for sustainable re-use.

Version 0.2

- Consistent use of 'snake_case' format for JSON field names
- Added missing lab identifiers
- Added default string 'coming' for missing required field values
"""

import xlrd
import glob

try:
    import simplejson as json
except ImportError:
    import json
    
# I/O routines
    
def getFileList(pattern):
    return glob.glob(pattern)

# managing JSON I/O
    
def getJSONPrettyDumpAsString(myDict):
    # returns a pretty formatted JSON fragment from a Python dictionary
    jsonDump = json.dumps(myDict, indent=4, separators=(', ', ': '))
    return jsonDump
    
def loadJSONFromFile(fileName):
    # returns JSON-file's content as Python dictionary
    with open(fileName) as f:
        jsonDict = json.load(f)
        return jsonDict

def writeJSONFile(fileName, jsonData):
    with open(fileName, 'w', encoding='utf-8') as f:
        json.dump(jsonData, f, ensure_ascii=False, indent=4)

# managing Excel workbooks

def getSheetNames(book):
    # returns a list of sheet names in given workbook
    sheetList = []
    for s in book.sheets():
        sheetList.append(s.name)
    return sheetList

def getSheet(book,sheetName):
    # returns a sheet object sheetname in workbook
    return book.sheet_by_name(sheetName)

def openWorkBook(path):
    # Open and read an Excel file
    book = xlrd.open_workbook(path)
    return book

# specific code for the EPOS TNA/Infrastructure portal project
    
def getCellNumberByStringValue(sheet,key):
    # assumption: keys are unique; no exception handling therefore at this moment
    cellNumber = []
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            if sheet.cell_value(row, col) == key:
                cellNumber = [row, col]
    return cellNumber

def getMultipleCellNumbersByStringValue(sheet,key):
    # keys are not necessarily unique; no exception handling therefore at this moment
    cellNumbers = []
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            if sheet.cell_value(row, col) == key:
                cellNumbers.append([row, col])
    return cellNumbers
    
def getFieldValue(sheet,keyName,templateSheet):
    # This routine is only to be used to get the general lab info; equipment information is differently structured in the Excel-template
    # pre-assumptions:
    # field names are unique in the template
    # value is to be found in the cell next to the right of the label-field cell
    # To remove fill in instruction values: value is checked against value from template: if equal return empty string
    cellNumber = getCellNumberByStringValue(sheet,keyName)
    if len(cellNumber) != 0:
        returnVal = sheet.cell(int(cellNumber[0]),int(cellNumber[1]+1)).value
    else:
        returnVal =''
    # check against default value from template:
    templateCellNumber = getCellNumberByStringValue(templateSheet,keyName)
    if len(templateCellNumber) != 0:
        templateVal = templateSheet.cell(int(templateCellNumber[0]),int(templateCellNumber[1]+1)).value
    else:
        templateVal =''
    if returnVal == templateVal:
        returnVal = ''
    return returnVal

def getFieldValueByAdjacentCellNum(sheet,cellNum):
    # pre-assumptions:
    # value is to be found in the cell next to the right
    returnVal = sheet.cell(int(cellNum[0]),int(cellNum[1]+1)).value
    return returnVal

def getFieldValueByCellNum(sheet,cellNum,templateDefaultValuesList):  
    try:
        returnVal = sheet.cell(int(cellNum[0]),int(cellNum[1])).value
        # skip default values from template
        if returnVal in templateDefaultValuesList:
            returnVal = ''
    except:
        returnVal = ''

    return returnVal

def getListOfMergedCellValues(sheet):
    # merged cells are separator rows in the template
    # returns a list values for these rows
    mergedCellsList = sheet.merged_cells
    mergedValues = []
    for c in mergedCellsList:
        mergedValues.append(sheet.cell(c[0],0).value)
    return mergedValues

def checkOnEmptyRow(sheet,nrow):
    empty = True
    row = sheet.row(nrow)
    for cell in row:
        if not (cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK)):
            empty = False
    return empty

def readLabIdentifier(labName):
    # we know there are occasionally double spaces encountered
    labId = ''
    if labName != '':
        labNameStripped = labName.replace('  ', ' ')
        labNameStripped = labNameStripped.replace('"', '')
        labs = loadJSONFromFile('sources/labs_unicode_updated.json')
        for lab in labs:
            name = lab["name"].replace('  ', ' ')
            name = name.replace('"', '')
            if name.find(labNameStripped) != -1:
                # when the name provided in the template is a non-empty substring then we have a hit
                labId = lab["id"]
    return labId
        

def fillGeneralLabInfo(sheet,domain,templateSheet):

    # address
    address = {'street_with_number' : getFieldValue(sheet,'Facility address (street + number)',templateSheet),
               'postal_code' : getFieldValue(sheet,'Facility address (postcode)',templateSheet),
               'city' : getFieldValue(sheet,'Facility address (city)',templateSheet),
               'country' : getFieldValue(sheet,'Facility address (country)',templateSheet)
               }
    
    # gps
    gps = {'gps_lat': getFieldValue(sheet, 'Facility gpsLat (decimal degree)',templateSheet),
           'gps_lon': getFieldValue(sheet, 'Facility gpsLon (decimal degree)',templateSheet)
           }
    
    
    # affiliation
    # Currently not further elaborated upon.
    # The Excel template does not currently request address and PIC e.g.
    affiliation = {'legal_name': getFieldValue(sheet, 'Affiliation of Facility contact person',templateSheet),
                   'identifier': {
                           'id_type': '',
                           'id_value': ''},
                   'address': {
                           'street_with_number': '',
                           'postal_code': '',
                           'city': '',
                           'country': ''}
                   }
      
    # contact person
    contact_person = {'first_name' : getFieldValue(sheet, 'Facility contact person (first name)',templateSheet),
                      'family_name': getFieldValue(sheet,'Facility contact person (family name)',templateSheet),
                      'identifier': {
                              'id_type': '',
                              'id_value': getFieldValue(sheet,'Facility contact person ID ',templateSheet)},
                      'email': getFieldValue(sheet, 'Email of Facility contact person',templateSheet),
                      'affiliation': affiliation}
    
    # facility
    labName = getFieldValue(sheet, 'Facility Name',templateSheet)
    if labName.lower() in ['other', '']:
        labName = getFieldValue(sheet, 'Facility Name (if other)',templateSheet)
    labName = labName.replace('  ', ' ') # remove possible double spaces
    labId = getFieldValue(sheet,'Facility ID',templateSheet)
    if labId in ['will be assigned later', '']:
        # if not already provided
        labId = readLabIdentifier(labName)
    # links to the MSL CKAN catalogue portal
    dataServices = [{'service_type': 'data_publications_access',
                     'link_label' : 'Go to data publications from this lab (MSL TCS catalogue portal)',
                     'URL': 'https://epos-msl.uu.nl/organization/' + labId},
                  {'service_type': 'data_publications_get',
                   'link_label': 'Retrieve data publications from this lab',
                   'URL': 'https://epos-msl.uu.nl/ics/api.php?Lab' + labId,
                   'payload' : 'json'}]

    if labId == '':
        log = {'Missing identifier for' : fileName + ' (labName = ' + labName + ')'}
        logIdentifiers.append(log)

    facility = {'facility' : {'type' : 'laboratory',
                'ID' : labId,
                'RI_name' : getFieldValue(sheet, 'RI name',templateSheet),
                'name' : labName,
                'general_description' : getFieldValue(sheet,'Lab information',templateSheet),
                'address' : address,
                'gps' : gps,
                'website' : getFieldValue(sheet, 'Facility website',templateSheet),
                'contact_person' : contact_person,
                'lab_services' : fillLabServices(domain, sheet,templateSheet),
                'data_services' : dataServices},
                }

    return facility


def fillLabServices(domain, sheet,templateSheet):
    returnValue = {}
    if domain == 'paleomag':
        returnValue = fillPaleoLabServices(sheet,templateSheet)
    else:
        if domain == 'analogue':
            returnValue = fillAnalogueLabServices(sheet,templateSheet)
        else:
            if domain == 'rock_physics':
                returnValue = fillRockPhysicsLabServices(sheet,templateSheet)
            else:
                if domain == 'analytical':
                    returnValue = fillAnalyticalLabServices(sheet, templateSheet)
    return returnValue
        
def fillPaleoLabServices(sheet,templateSheet):
    # equipment types
    equipmentTypes = []
    # find header row
    equipmentTypeHeaderRowNum = getCellNumberByStringValue(sheet, 'Equipment type')[0]
    valueRow = equipmentTypeHeaderRowNum+1
    equipmentTemplateTypeHeaderRowNum = getCellNumberByStringValue(templateSheet, 'Equipment type')[0]
    valueTemplateRow = equipmentTemplateTypeHeaderRowNum+1
    
    # in the template valueRow contains possible default values against which we must check:
    defaults = []
    for x in range(0, 11):
        defaults.append(getFieldValueByCellNum(templateSheet,(valueTemplateRow,x),[]))
    
    # get separator by checking on merged cells values
    mergedCells = getListOfMergedCellValues(sheet)
    mergedCells.append('') # otherwise we'll fail on the empty row

    eqType = getFieldValueByCellNum(sheet,(valueRow,0),defaults)

    while (not eqType in mergedCells) and (valueRow < sheet.nrows):
    #while eqType != '':
        if not checkOnEmptyRow(sheet,valueRow):    
            equipment_name = getFieldValueByCellNum(sheet,(valueRow,2),defaults)
            equipment_custom_name = getFieldValueByCellNum(sheet,(valueRow,3),defaults)
            if equipment_custom_name != '':
                # turn base name into group, assign 'other value' to name
                entrancePart1 = {'equipment_type' : eqType,
                                 'equipment_group' : getFieldValueByCellNum(sheet,(valueRow,1),defaults),
                                 'equipment_name' : equipment_name,
                                 'equipment_secondary_name': equipment_custom_name}
                log = {'domain' : 'paleomag',
                       'category' : 'equipment',
                       'File' : fileName,
                       'base_name' : equipment_name,
                       'specific name' : equipment_custom_name}
                logInfo.append(log)
            
            else:
                entrancePart1 = {'equipment_type' : eqType,
                                 'equipment_group' : getFieldValueByCellNum(sheet,(valueRow,1),defaults),
                                 'equipment_name' : equipment_name}
            if eqType == '':
                log = {'empty equipment type in non-empty entrance' : fileName}
                logInfo.append(log)
            entrancePart2 = {'equipment_brand' : getFieldValueByCellNum(sheet,(valueRow,4),defaults),
                        'equipment_contact_person' : getFieldValueByCellNum(sheet,(valueRow,5),defaults),
                        'email_equipment_contact_person' : getFieldValueByCellNum(sheet,(valueRow,6),defaults),
                        'equipment_website' : getFieldValueByCellNum(sheet,(valueRow,7),defaults),
                        'equipment_specifics_and_comments' : getFieldValueByCellNum(sheet,(valueRow,8),defaults),
                        'equipment_quantity' : getFieldValueByCellNum(sheet,(valueRow,9),defaults),
                        'references' : getFieldValueByCellNum(sheet,(valueRow,10),defaults)}
            entrance = {}
            entrance.update(entrancePart1)
            entrance.update(entrancePart2)
            equipmentTypes.append(entrance)
            equipmentTypes.append(entrance)
        valueRow += 1
        eqType = getFieldValueByCellNum(sheet,(valueRow,0),defaults)
        
    # measurement types
    measurementTypes = []
    # find header row
    measurementTypeHeaderRowNum = getCellNumberByStringValue(sheet, 'Measurement type')[0]
    valueRow = measurementTypeHeaderRowNum+1
    measurementTypeTemplateHeaderRowNum = getCellNumberByStringValue(templateSheet, 'Measurement type')[0]
    templateValueRow = measurementTypeTemplateHeaderRowNum+1
    defaults.clear()
    for x in range(6):
        defaults.append(getFieldValueByCellNum(templateSheet,(templateValueRow,x),['']))
    
    measurementType = getFieldValueByCellNum(sheet,(valueRow,0),defaults)

    while (not measurementType in mergedCells) and (valueRow < sheet.nrows):
        measurement_name = getFieldValueByCellNum(sheet,(valueRow,2),defaults)
        if measurement_name.lower() == 'other':
            measurement_custom_name = getFieldValueByCellNum(sheet,(valueRow,3),defaults)
            if measurement_custom_name != '':
                measurement_name = measurement_custom_name
        entrance = {'measurement_type' : measurementType,
                    'measurement_group' : getFieldValueByCellNum(sheet,(valueRow,1),defaults),
                    'measurement_name' : measurement_name,
                    'measured_type_specifics_and_comments' : getFieldValueByCellNum(sheet,(valueRow,4),defaults),
                    'references' : getFieldValueByCellNum(sheet,(valueRow,5),defaults)}
        measurementTypes.append(entrance)
        valueRow += 1
        measurementType = getFieldValueByCellNum(sheet,(valueRow,0),defaults)

        
    returnValue = {'research_field': 'Paleomagnetism', 
                      'EPOS_subdomain': ['Paleomagnetic and magnetic data'],
                      'equipment' : equipmentTypes,
                      'measurement' : measurementTypes}
    return returnValue

def fillAnalogueLabServices(sheet,templateSheet):
    # equipment types
    equipmentTypes = []
    # find header row
    equipmentTypeHeaderRowNum = getCellNumberByStringValue(sheet, 'Equipment type')[0]
    # set the starting row
    valueRow = equipmentTypeHeaderRowNum+1
    # find applicable header row in template
    equipmentTemplateTypeHeaderRowNum = getCellNumberByStringValue(templateSheet, 'Equipment type')[0]
    # set the first template entrance row that contains fixed information strings that have to skipped
    valueTemplateRow = equipmentTemplateTypeHeaderRowNum+1    
    # get the possible default information string values against which we must check:
    defaults = []
    for x in range(0, 11):
        defaults.append(getFieldValueByCellNum(templateSheet,(valueTemplateRow,x),[]))
    # get category separator by checking on merged cells values; merged cells are inter-category rows
    mergedCells = getListOfMergedCellValues(sheet)
    mergedCells.remove('') # TECMOD has merged empty line

    #start from first entrance row for equipment category
    eqType = getFieldValueByCellNum(sheet,(valueRow,0),defaults)

    while (not eqType in mergedCells) and (valueRow < sheet.nrows):
        # lastrow check for safety: lab managers could have deleted the other categories
        # we have to be aware of empty rows! (see e.g. TECMOD analogue lab rance) 
        if not checkOnEmptyRow(sheet,valueRow):
            equipment_name = getFieldValueByCellNum(sheet,(valueRow,1),defaults)
            # check whether the 'other equipment name' column was filled
            secondaryName = getFieldValueByCellNum(sheet,(valueRow,2),defaults)
            if secondaryName != '':
                # turn base name into group, assign 'other value' to name
                entrancePart1 = {'equipment_type' : eqType,
                                 'equipment_group' : equipment_name,
                                 'equipment_name' : secondaryName}
                log = {'domain' : 'analogue',
                       'category' : 'equipment',
                       'File' : fileName,
                       'base_name' : equipment_name,
                       'specific name' : secondaryName}
                logInfo.append(log)
            
            else:
                entrancePart1 = {'equipment_type' : eqType,
                             'equipment_name' : equipment_name}
            if eqType == '':
                log = {'empty equipment type in non-empty entrance' : fileName}
                logInfo.append(log)
            # common part:
            entrancePart2 = {'equipment_brand' : getFieldValueByCellNum(sheet,(valueRow,3),defaults),
                             'equipment_contact_person' : getFieldValueByCellNum(sheet,(valueRow,4),defaults),
                             'email_equipment_contact_person' : getFieldValueByCellNum(sheet,(valueRow,5),defaults),
                             'equipment_website' : getFieldValueByCellNum(sheet,(valueRow,6),defaults),
                             'equipment_specifics_and_comments' : getFieldValueByCellNum(sheet,(valueRow,7),defaults),
                             'equipment_quantity' : getFieldValueByCellNum(sheet,(valueRow,8),defaults),
                             'references' : getFieldValueByCellNum(sheet,(valueRow,9),defaults)}
            entrance = {}
            entrance.update(entrancePart1)
            entrance.update(entrancePart2)
            equipmentTypes.append(entrance)
        valueRow += 1
        eqType = getFieldValueByCellNum(sheet,(valueRow,0),defaults)
    
    # material
    materialTypes = []
    # find header row
    materialTypeHeaderRowNum = getCellNumberByStringValue(sheet, 'Material')[0]
    valueRow = materialTypeHeaderRowNum+1
    materialTypeTemplateHeaderRowNum = getCellNumberByStringValue(templateSheet, 'Material')[0]
    templateValueRow = materialTypeTemplateHeaderRowNum+1
    defaults.clear()
    for x in range(5):
        defaults.append(getFieldValueByCellNum(templateSheet,(templateValueRow,x),['']))
    
    materialType = getFieldValueByCellNum(sheet,(valueRow,0),defaults)

    while (not (materialType in mergedCells)) and (valueRow < sheet.nrows):
        if not checkOnEmptyRow(sheet,valueRow):
            # check whether the 'other material type' column was filled
            material_custom_name = getFieldValueByCellNum(sheet,(valueRow,1),defaults)
            if material_custom_name != '':
                entrancePart1 = {'material_group' : materialType,
                                 'material' : material_custom_name}
                log = {'domain' : 'analogue',
                       'category' : 'Material',
                       'File' : fileName,
                       'base_name' : materialType,
                       'specific name' : material_custom_name}
                logInfo.append(log)
                
            else:
                entrancePart1 = {'material' : materialType}
    
            entrancePart2 = {'material_brand' : getFieldValueByCellNum(sheet,(valueRow,2),defaults),
                        'material_specifics_and_comments' : getFieldValueByCellNum(sheet,(valueRow,3),defaults),
                        'references' : getFieldValueByCellNum(sheet,(valueRow,4),defaults)}
            entrance = {}
            entrance.update(entrancePart1)
            entrance.update(entrancePart2)
            materialTypes.append(entrance)
        valueRow += 1
        materialType = getFieldValueByCellNum(sheet,(valueRow,0),defaults)


    # Measured property
    measuredProperties = []
    # find header row
    propertyHeaderRowNum = getCellNumberByStringValue(sheet, 'Measured property')[0]
    valueRow = propertyHeaderRowNum+1
    propertyTemplateHeaderRowNum = getCellNumberByStringValue(templateSheet, 'Measured property')[0]
    templateValueRow = propertyTemplateHeaderRowNum+1
    defaults.clear()
    for x in range(4):
        defaults.append(getFieldValueByCellNum(templateSheet,(templateValueRow,x),['']))
    
    propertyType = getFieldValueByCellNum(sheet,(valueRow,0),defaults)

    while (not propertyType in mergedCells) and (valueRow < sheet.nrows):
        if not checkOnEmptyRow(sheet,valueRow):
            property_custom_name = getFieldValueByCellNum(sheet,(valueRow,1),defaults)
            if property_custom_name != '':
                entrancePart1 = {'measured_property_group' : propertyType,
                                 'measured_property' : property_custom_name}
                log = {'domain' : 'analogue',
                       'category' : 'Measured property',
                       'File' : fileName,
                       'base_name' : propertyType,
                       'specific name' : property_custom_name}
                logInfo.append(log)
                
            else:
                entrancePart1 = {'measured_property' : propertyType}
            entrancePart2 = {'measured_property_specifics_and_comments' : getFieldValueByCellNum(sheet,(valueRow,2),defaults),
                        'references' : getFieldValueByCellNum(sheet,(valueRow,3),defaults)}
            entrance = {}
            entrance.update(entrancePart1)
            entrance.update(entrancePart2)
            measuredProperties.append(entrance)
        valueRow += 1
        if valueRow <= sheet.nrows:
            # this was not the last row; we can safely read beyond
            propertyType = getFieldValueByCellNum(sheet,(valueRow,0),defaults)

        
    returnValue = {'research_field': 'Analogue modelling of geologic processes', 
                      'EPOS_subdomain': ['Analogue models on tectonic processes'],
                      'equipment' : equipmentTypes,
                      'material' : materialTypes,
                      'measured_property' : measuredProperties}
    return returnValue

def fillRockPhysicsLabServices(sheet,templateSheet):
    # equipment types
    equipmentTypes = []
    # find header row
    equipmentTypeHeaderRowNum = getCellNumberByStringValue(sheet, 'Equipment type')[0]
    valueRow = equipmentTypeHeaderRowNum+1
    equipmentTemplateTypeHeaderRowNum = getCellNumberByStringValue(templateSheet, 'Equipment type')[0]
    valueTemplateRow = equipmentTemplateTypeHeaderRowNum+1
    
    # in the template valueRow contains possible default values against which we must check:
    defaults = []
    for x in range(0, 10):
        defaults.append(getFieldValueByCellNum(templateSheet,(valueTemplateRow,x),[]))
    
    # get separator by checking on merged cells values
    mergedCells = getListOfMergedCellValues(sheet)
    mergedCells.remove('') # there could be merged cells with empty string values and we need to allow them in between

    eqType = getFieldValueByCellNum(sheet,(valueRow,0),defaults)

    while (not eqType in mergedCells) and (valueRow < sheet.nrows):
    #while eqType != '':
        if not checkOnEmptyRow(sheet,valueRow):
            equipment_name = getFieldValueByCellNum(sheet,(valueRow,1),defaults)
            secondaryName = getFieldValueByCellNum(sheet,(valueRow,2),defaults)
            if secondaryName != '':
                entrancePart1 = {'equipment_type' : eqType,
                            'equipment_group' : equipment_name,
                            'equipment_name' : secondaryName}
                log = {'domain': 'rock physics',
                       'category' : 'equipment',
                       'File' : fileName,
                       'base_name' : equipment_name,
                       'specific name' : secondaryName}
                logInfo.append(log)
            else:
                entrancePart1 = {'equipment_type' : eqType,
                            'equipment_name' : equipment_name}
            if eqType == '':
                log = {'empty equipment type in non-empty entrance' : fileName}
                logInfo.append(log)
            entrancePart2 = {'equipment_brand' : getFieldValueByCellNum(sheet,(valueRow,3),defaults),
                            'equipment_contact_person' : getFieldValueByCellNum(sheet,(valueRow,4),defaults),
                            'email_equipment_contact_person' : getFieldValueByCellNum(sheet,(valueRow,5),defaults),
                            'equipment_website' : getFieldValueByCellNum(sheet,(valueRow,6),defaults),
                            'equipment_specifics_and_comments' : getFieldValueByCellNum(sheet,(valueRow,7),defaults),
                            'equipment_quantity' : getFieldValueByCellNum(sheet,(valueRow,8),defaults),
                            'references' : getFieldValueByCellNum(sheet,(valueRow,9),defaults)}
            entrance = {}
            entrance.update(entrancePart1)
            entrance.update(entrancePart2)
            equipmentTypes.append(entrance)
        valueRow += 1
        eqType = getFieldValueByCellNum(sheet,(valueRow,0),defaults)
        
    returnValue = {'research_field': 'Rock/melt physics & Microscopy', 
                      'EPOS_subdomain': ['Rock/melt physics & Microscopy'],
                      'equipment' : equipmentTypes}
    return returnValue

def fillAnalyticalLabServices(sheet,templateSheet):
    # equipment types
    equipmentTypes = []
    # find header row
    equipmentTypeHeaderRowNum = getCellNumberByStringValue(sheet, 'Equipment name')[0]
    valueRow = equipmentTypeHeaderRowNum+1
    equipmentTemplateTypeHeaderRowNum = getCellNumberByStringValue(templateSheet, 'Equipment name')[0]
    valueTemplateRow = equipmentTemplateTypeHeaderRowNum+1
    
    # in the template valueRow contains possible default values against which we must check:
    defaults = []
    for x in range(0, 9):
        defaults.append(getFieldValueByCellNum(templateSheet,(valueTemplateRow,x),[]))
    
    # get separator by checking on merged cells values
    mergedCells = getListOfMergedCellValues(sheet)
    mergedCells.append('') # otherwise we'll fail on the empty row

    eqType = getFieldValueByCellNum(sheet,(valueRow,0),defaults)

    while not (eqType in mergedCells):
    #while eqType != '':
    
        equipment_name = getFieldValueByCellNum(sheet,(valueRow,0),defaults)
        secondaryName = getFieldValueByCellNum(sheet,(valueRow,1),defaults)
        if secondaryName != '':
            entrancePart1 = {'equipment_group' : equipment_name,
                             'equipment_name' : secondaryName}
            log = {'domain' : 'analytical',
                   'category' : 'equipment',
                   'File' : fileName,
                   'base_name' : equipment_name,
                   'specific name' : secondaryName}
            logInfo.append(log)
            
        else:
            entrancePart1 = {'equipment_name' : equipment_name}

        entrancePart2 = {'equipment_brand' : getFieldValueByCellNum(sheet,(valueRow,2),defaults),
                    'equipment_contact_person' : getFieldValueByCellNum(sheet,(valueRow,3),defaults),
                    'email_equipment_contact_person' : getFieldValueByCellNum(sheet,(valueRow,4),defaults),
                    'equipment_website' : getFieldValueByCellNum(sheet,(valueRow,5),defaults),
                    'equipment_specifics_and_comments' : getFieldValueByCellNum(sheet,(valueRow,6),defaults),
                    'equipment_quantity' : getFieldValueByCellNum(sheet,(valueRow,7),defaults),
                    'references' : getFieldValueByCellNum(sheet,(valueRow,8),defaults)}
        entrance = {}
        entrance.update(entrancePart1)
        entrance.update(entrancePart2)
        equipmentTypes.append(entrance)
        valueRow += 1
        eqType = getFieldValueByCellNum(sheet,(valueRow,0),defaults)
        
    returnValue = {'research_field': 'Solid Earth Geochemistry', 
                      'EPOS_subdomain': ['Geochemical data (elemental and isotope geochemistry)'],
                      'equipment' : equipmentTypes}
    return returnValue
    


def processExcel(sourceDomain,fullPath, outputDir,templateFile):
    fileName = fullPath.rsplit('/',1)[1]
    #to be sure that we have no space in the filename:
    fileName = fileName.replace(' ', '_')
    book = openWorkBook(fullPath)
    sheet = book.sheet_by_index(0) #because rock phys uu sheetname gives an unexplained error
    template = openWorkBook(templateFile)
    templateSheet = template.sheet_by_index(0)
    labInfo = fillGeneralLabInfo(sheet,sourceDomain,templateSheet)
    if labInfo['facility'].get('ID') != '':
        # we only add labs with a valid identifier to the allLabs export for ICS
        allLabsExport.append(labInfo)
    writeJSONFile(outputDir + '/' + fileName + '.json',labInfo)

#---------------
    
if __name__ == "__main__":
    # domain: [category,sources path, template file]]
    domains = [['analogue','sources/analogue/*.xlsx','sources/analogue/template/TEMPLATE_Laboratory_description_analoguemodelling.xlsx'],
               ['analytical','sources/analytical/*.xlsx','sources/analytical/template/Laboratory description_analytical&microscopy labs_TEMPLATE.xlsx'],
               ['paleomag','sources/paleomag/*.xlsx','sources/paleomag/template/TEMPLATE_Laboratory description_paleomagnetism_V5.xlsx'],
               ['rock_physics','sources/rock_physics/*.xlsx','sources/rock_physics/template/Laboratory description_rock physics_TEMPLATE.xlsx']]
    
    # managing 'other' entrances' with globals
    global logInfo
    global fileName
    global logIdentifiers
    logInfo = []
    fileName = ''
    logIdentifiers = []
    
    # for creating one single exportfile
    global allLabsExport
    allLabsExport = []
    
    #global numMissingIdentifiers
    #numMissingIdentifiers = 0
    
    for domain in domains:
        domainName = domain[0]
        sourceFiles = domain[1]
        templateFile = domain[2]
        files = getFileList(sourceFiles)
        outputDir = sourceFiles.rsplit('/',1)[0] + '/json_out'
        for file in files:
            fileName = file.rsplit('/',1)[1]
            processExcel(domainName,file, outputDir,templateFile)
            
    log = []
    log.append(logIdentifiers)
    log.append(logInfo)
    writeJSONFile('sources/log.json',log)
    infraStructures = {'infrastructures' : allLabsExport}
    writeJSONFile('sources/allLabs.json', infraStructures)
 
    
    #print(testJSONFileIO('sources/lab_info_general.json'))
    
#---------------

