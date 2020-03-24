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
from infra_msl.msl_portal import tcs_portal

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
    cellNumber = [-1,-1] # returned when value is not found
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

def readLabIdentifier(labName, IDsFile):
    # we know there are occasionally double spaces encountered
    labId = ''
    if labName != '':
        print(labName)
        labNameStripped = labName.replace('  ', ' ')
        labNameStripped = labNameStripped.replace('"', '')
        labNameStripped = labNameStripped.replace(',', '')
        labs = loadJSONFromFile(IDsFile)
        for lab in labs:
            name = lab["name"].replace('  ', ' ')
            name = name.replace('"', '')
            name = name.replace(',', '')
            if name.find(labNameStripped) != -1:
                # when the name provided in the template is a non-empty substring then we have a hit
                labId = lab["id"]
    return labId
        

def fillGeneralLabInfo(sheet,research_field, subdomain,templateSheet,IDsFile):

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
    # we use representation names from V6.1 onwards:
    labName = getFieldValue(sheet, 'Facility Name (if other)',templateSheet)
    
    labNameForIdRetrieval = getFieldValue(sheet, 'Facility Name',templateSheet)
    if labNameForIdRetrieval.lower() in ['other', '']:
        labNameForIdRetrieval = getFieldValue(sheet, 'Facility Name (if other)',templateSheet)
    labName = labName.replace('  ', ' ') # remove possible double spaces
    labNameForIdRetrieval = labNameForIdRetrieval.replace('  ', ' ') # remove possible double spaces
    labId = getFieldValue(sheet,'Facility ID',templateSheet)
    if labId in ['will be assigned later', '']:
        # if not already provided
        labId = readLabIdentifier(labNameForIdRetrieval,IDsFile)
    # links to the MSL CKAN catalogue portal
    # We use the TCS Portal Webservice to check whether this lab has data publications already:
    
    numOfPublications = tcs_portal.retrieveNumberOfLabPublications(labId)
    if numOfPublications > 0:
        dataServices = [{'service_type': 'data_publications_access',
                         'link_label' : 'Go to data publications from this lab (MSL TCS catalogue portal)',
                         'URL': 'https://epos-msl.uu.nl/organization/' + labId},
                      {'service_type': 'data_publications_get',
                       'link_label': 'Retrieve data publications from this lab',
                       'URL': 'https://epos-msl.uu.nl/ics/api.php?Lab=' + labId,
                       'payload' : 'json'}]
    else:
        dataServices = []
    
    dataServices.append({'service_type': 'TCS_portal_redirection',
                    'link_label': 'More facility information',
                    'URL': 'https://epos-msl.uu.nl/organization/about/' + labId})


    if labId == '':
        log = {'Missing identifier for' : fileName + ' (labName = ' + labName + ')'}
        logIdentifiers.append(log)

    facility = {'facility' : {'type' : 'laboratory',
                'lab_id' : labId,
                'research_infrastructure_name' : getFieldValue(sheet, 'RI name',templateSheet),
                'facility_name' : labName,
                'address' : address,
                'gps' : gps,
                'website' : getFieldValue(sheet, 'Facility website',templateSheet),
                'lab_services' : fillLabServices(research_field, subdomain, sheet,templateSheet),
                'data_services' : dataServices},
                }
        #'contact_person' : contact_person,
        # 'general_description' : getFieldValue(sheet,'Lab information',templateSheet),

    return facility

def fillLabServices(research_field, subdomain, sheet,templateSheet):
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
                        'equipment_website' : getFieldValueByCellNum(sheet,(valueRow,7),defaults),
                        'equipment_specifics_and_comments' : getFieldValueByCellNum(sheet,(valueRow,8),defaults),
                        'equipment_quantity' : getFieldValueByCellNum(sheet,(valueRow,9),defaults),
                        'references' : getFieldValueByCellNum(sheet,(valueRow,10),defaults)}
            
            # The next fields are not exported as of 2019-09-13. We have to check against GPRD
            # 'equipment_contact_person' : getFieldValueByCellNum(sheet,(valueRow,5),defaults),
            # 'email_equipment_contact_person' : getFieldValueByCellNum(sheet,(valueRow,6),defaults),
           
            entrance = {}
            entrance.update(entrancePart1)
            entrance.update(entrancePart2)
            equipmentTypes.append(entrance)
            
        valueRow += 1
        eqType = getFieldValueByCellNum(sheet,(valueRow,0),defaults)
        
    # We'll sort on equipment names: discussed in Prague Sept 2019
    from operator import itemgetter
    newList = sorted(equipmentTypes, key=itemgetter('equipment_name'))
    equipmentTypes = newList
        
    # measurement types
    measurementTypes = []
    # find header row
    # update 2019-09-13: not all domains provide this, so we have to build in an exception
    
    measurementTypeHeaderRowNum = getCellNumberByStringValue(sheet, 'Measurement type')[0]
    
    if measurementTypeHeaderRowNum != -1:
    
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
            
        # Sorting
        #from operator import itemgetter
        newList = sorted(measurementTypes, key=itemgetter('measurement_name'))
        measurementTypes = newList
        
    else:
        measurementTypes ={}
    
    returnValue = {'research_field': research_field, 
                      'subdomain': [subdomain],
                      'equipment' : equipmentTypes,
                      'measurement' : measurementTypes}
    
    
    return returnValue

   


def processExcel(research_field, subdomain,fullPath, outputDir,templateFile,IDsFile):
    fileName = fullPath.rsplit('/',1)[1]
    #to be sure that we have no space in the filename:
    fileName = fileName.replace(' ', '_')
    book = openWorkBook(fullPath)
    sheet = book.sheet_by_index(0) #because rock phys uu sheetname gives an unexplained error
    template = openWorkBook(templateFile)
    templateSheet = template.sheet_by_index(0)
    labInfo = fillGeneralLabInfo(sheet,research_field, subdomain,templateSheet,IDsFile)
    if labInfo['facility'].get('lab_id') != '':
        # we only add labs with a valid identifier to the allLabs export for ICS
        allLabsExport.append(labInfo)
    writeJSONFile(outputDir + '/' + fileName + '.json',labInfo)

#---------------
    
def getSources(root):
    import glob
    sourceFiles = glob.glob(root + '[A-Z,a-z]*.xlsx')
    for file in sourceFiles:
        if file.find('TEMPLATE') > -1:
            sourceFiles.remove(file)
    return sourceFiles

# running the conversion script with special parameters from the jupyter notebook:

def runConversion(PALEOTEMPLATE,PALEOFILES,ROCKTEMPLATE,ROCKFILES,IDS_FILE,JSONOUTPUT):
    domains = [{'research_field': 'Paleomagnetism',
                'subdomain': 'Paleomagnetic and magnetic data',
                'sourceDir' : PALEOFILES,
                'sourceFiles': getSources(PALEOFILES),
                'template': PALEOTEMPLATE},
        {'research_field': 'Rock physics',
                'subdomain': 'Rock and melt physical properties',
                'sourceDir' : ROCKFILES,
                'sourceFiles': getSources(ROCKFILES),
                'template': PALEOTEMPLATE}]
    
    IDENTIFIERSFILE = IDS_FILE
    
    
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
    
    import os
    
    #global numMissingIdentifiers
    #numMissingIdentifiers = 0
    
    for domain in domains:
        research_field = domain['research_field']
        subdomain = domain['subdomain']
        files = domain['sourceFiles']
        templateFile = domain['template']
        outputDir = domain['sourceDir'] + '/json_out'
        CHECK_FOLDER = os.path.isdir(outputDir)
        if not CHECK_FOLDER:
            os.makedirs(outputDir)
            print("created folder : ", outputDir)   
        else:
            print(outputDir, "outputfolder exists.")
        for file in files:
            fileName = file.rsplit('/',1)[1]
            processExcel(research_field, subdomain,file, outputDir,templateFile,IDENTIFIERSFILE)
            
    log = []
    log.append(logIdentifiers)
    log.append(logInfo)
    # writeJSONFile(outputDir + '/log.json',log)
    infraStructures = {'infrastructures' : allLabsExport}
    # writeJSONFile(outputDir + '/allpaleomagLabs_v6.1.json', infraStructures)
    # and write the new source for the service:
    writeJSONFile(JSONOUTPUT, infraStructures)




#---------------
    
# running the whole script with standard parameters:
    
if __name__ == "__main__":
    
    PALEOTEMPLATE = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/Paleomag/Updated_Paleomag_Lab_Description_V6.1/TEMPLATE_Laboratory description_paleomagnetism_V6.1.xlsx'
    PALEOFILES = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/Paleomag/Updated_Paleomag_Lab_Description_V6.1/'
    
    ROCKTEMPLATE = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/Rock physics/Updated_Rock_Physics_Lab_Description_V3/TEMPLATE_Laboratory description_rock physics_V3.xlsx'
    ROCKFILES = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/Rock physics/Updated_Rock_Physics_Lab_Description_V3/'
    
    IDENTIFIERSFILE = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/lab_identifiers.json'

    
    domains = [{'research_field': 'Paleomagnetism',
                'subdomain': 'Paleomagnetic and magnetic data',
                'sourceDir' : PALEOFILES,
                'sourceFiles': getSources(PALEOFILES),
                'template': PALEOTEMPLATE},
        {'research_field': 'Rock physics',
                'subdomain': 'Rock and melt physical properties',
                'sourceDir' : ROCKFILES,
                'sourceFiles': getSources(ROCKFILES),
                'template': PALEOTEMPLATE}]
    
    # Analogue modelling of geologic processes
    # Analytical and microscopy data
    
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
    
    import os
    
    #global numMissingIdentifiers
    #numMissingIdentifiers = 0
    
    for domain in domains:
        research_field = domain['research_field']
        subdomain = domain['subdomain']
        files = domain['sourceFiles']
        templateFile = domain['template']
        outputDir = domain['sourceDir'] + '/json_out'
        CHECK_FOLDER = os.path.isdir(outputDir)
        if not CHECK_FOLDER:
            os.makedirs(outputDir)
            print("created folder : ", outputDir)   
        else:
            print(outputDir, "outputfolder exists.")
        for file in files:
            fileName = file.rsplit('/',1)[1]
            processExcel(research_field, subdomain,file, outputDir,templateFile,IDENTIFIERSFILE)
            
    log = []
    log.append(logIdentifiers)
    log.append(logInfo)
    # writeJSONFile(outputDir + '/log.json',log)
    infraStructures = {'infrastructures' : allLabsExport}
    # writeJSONFile(outputDir + '/allpaleomagLabs_v6.1.json', infraStructures)
    # and write the new source for the service:
    writeJSONFile('/Users/otto/Documents/GitLab/inframsl/labs.json', infraStructures)
