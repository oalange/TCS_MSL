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
import tcs_portal
import hashlib

try:
    import simplejson as json
except ImportError:
    import json
    

#***************    
# I/O routines *
#***************
    
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

#***************************
# managing Excel workbooks *
#***************************
        
class excelWB:
    def __init__(self, pathToExcelFile, activeSheetName = ''):
        self.workbook = self.openWorkBook(pathToExcelFile)
        self.sheetnamesList = self.getSheetNames()
        if activeSheetName == '':
            self.activeSheet = self.getSheet(self.sheetnamesList[0]) # we initialize on first sheet
        else:
            self.activeSheet = self.getSheet(activeSheetName)

    def getSheetNames(self):
        # returns a list of sheet names in workbook
        sheetList = []
        for s in self.workbook.sheets():
            sheetList.append(s.name)
        return sheetList
    
    def getSheet(self,sheetName):
        # returns a sheet object sheetname in workbook
        return self.workbook.sheet_by_name(sheetName)
    
    def openWorkBook(self, path):
        # Open and read an Excel file
        book = xlrd.open_workbook(path)
        return book
    
    def getCellNumberByStringValue(self,key): #???
        # returns the first cell with value == key
        cellNumber = [-1,-1] # returned when value is not found
        
        for row in range(self.activeSheet.nrows):
            for col in range(self.activeSheet.ncols):
                if self.activeSheet.cell_value(row, col) == key:
                    cellNumber = [row, col]
        return cellNumber

#        row = 1
#        col = 1
#        while row < self.activeSheet.nrows:
#            while col < self.activeSheet.ncols:
#                if self.activeSheet.cell_value(row, col) == key:
#                    cellNumber = [row, col]
#                    row = self.activeSheet.nrows
#                    col = self.activeSheet.ncols
#                else:
#                    col += 1
#            row += 1
#        return cellNumber
    
    def getMultipleCellNumbersByStringValue(self,key):
        # keys are not necessarily unique, so we return a list
        cellNumbers = []
        for row in range(self.activeSheet.nrows):
            for col in range(self.activeSheet.ncols):
                if self.activeSheet.cell_value(row, col) == key:
                    cellNumbers.append([row, col])
        return cellNumbers
        
    def getFieldValue(self, keyName,templateSheet): #???
        # This routine is only to be used to get the general lab info; equipment information is differently structured in the Excel-template
        # pre-assumptions:
        # field names are unique in the template
        # value is to be found in the cell next to the right of the label-field cell
        # To remove fill in instruction values: value is checked against value from template: if equal return empty string
        cellNumber = self.getCellNumberByStringValue(keyName)
        if len(cellNumber) != 0:
            returnVal = self.activeSheet.cell(int(cellNumber[0]),int(cellNumber[1]+1)).value
        else:
            returnVal =''
        # check against default value from template:
        templateCellNumber = self.getCellNumberByStringValue(keyName)
        if len(templateCellNumber) != 0:
            templateVal = templateSheet.cell(int(templateCellNumber[0]),int(templateCellNumber[1]+1)).value
        else:
            templateVal =''
        if returnVal == templateVal:
            returnVal = ''
        return returnVal
    
    def getFieldValueByAdjacentCellNum(self,cellNum):
        # pre-assumptions:
        # value is to be found in the cell next to the right
        returnVal = self.activeSheet.cell(int(cellNum[0]),int(cellNum[1]+1)).value
        return returnVal
    
    def getFieldValueByCellNum(self,cellNum,templateDefaultValuesList):  
        try:
            returnVal = self.activeSheet.cell(int(cellNum[0]),int(cellNum[1])).value
            # skip default values from template
            if returnVal in templateDefaultValuesList:
                returnVal = ''
        except:
            returnVal = ''
    
        return returnVal
    
    def getListOfMergedCellValues(self):
        # merged cells are separator rows in the template
        # returns a list values for these rows
        mergedCellsList = self.activeSheet.merged_cells
        mergedValues = []
        for c in mergedCellsList:
            mergedValues.append(self.activeSheet.cell(c[0],0).value)
        return mergedValues
    
    def checkOnEmptyRow(self,nrow):
        empty = True
        row = self.activeSheet.row(nrow)
        for cell in row:
            if not (cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK)):
                empty = False
        return empty

# *************************************************
# class for managing identifiers for laboratories *
# *************************************************
    
class labIDs:
 # !!!
    def __init__(self, IDsFile):
        self.dictIDs = {'allIDs': []} # all template-based generated IDs
        self.matchGenIDs = {'generated_ids':[]} # all template-based generated IDs that match IDsFile
        self.unmatchGenIDs = {'generated_ids':[]} # all template-based generated IDs that do not match IDsFile
        self.IDsFile = IDsFile
        self.labs = loadJSONFromFile(self.IDsFile)


    def readLabIdentifier(self, labNames): # labNames is now a list of possible candidate strings
        # we know there are occasionally double spaces encountered
        labId = ''
        NotFound = True
        for labName in labNames:
            if labName != '' and NotFound:
                print(labName)
                labNameStripped = labName.replace('  ', ' ')
                labNameStripped = labNameStripped.replace('"', '')
                labNameStripped = labNameStripped.replace(',', '')
                labNameStripped = labNameStripped.lower()
                
                for lab in self.labs:
                    name = lab["inputstring"].replace('  ', ' ')
                    name = name.replace('"', '')
                    name = name.replace(',', '')
                    name = name.lower()
                    if name.find(labNameStripped) != -1:
                        # when the name provided in the template is a non-empty substring then we have a hit
                        labId = lab["id"]
                        NotFound = False
                        
                if NotFound: # we do a check on the labname field
                    for lab in self.labs:
                        name = lab["labname"].replace('  ', ' ')
                        name = name.replace('"', '')
                        name = name.replace(',', '')
                        name = name.lower()
                        if name.find(labNameStripped) != -1:
                            # when the name provided in the template is a non-empty substring then we have a hit
                            labId = lab["id"]
                            NotFound = False
                        
        return labId
    
    def generateLabIdentifier(self, labName):
        idBaseName = labName
        idBaseName = idBaseName.replace(' ', '')
        idBaseName = idBaseName.lower()
        return hashlib.md5(idBaseName.encode()).hexdigest()
    
    def addLabIdentifiers(self, idFile):
        # expects a valid identifiersFile and fills/updates the id-fields
        output = []
        entries = loadJSONFromFile(idFile)
        for entry in entries:
            updatedEntry = entry
            updatedEntry.update({'id': self.generateLabIdentifier(entry['inputstring'])})
            output.append(updatedEntry)
        writeJSONFile('out_' + idFile, output)
            
        
    
class lab:
    
    def __init__(self, labWb, research_field, subdomain, templateWb, IDs):
        self.ckan = tcs_portal.TCS_PortalRequests()
        self.labWb = labWb
        self.sheet = labWb.activeSheet
        self.research_field = research_field
        self.subdomain = subdomain
        self.templateWb = templateWb
        self.IDs = IDs
        self.labInfo = {}
        self.readLabInfo()
        self.labID = ''

    def fillGeneralLabInfo(self):
    
        # address
        facility_address = self.labWb.getFieldValue('Facility address (street + number)',self.templateWb.activeSheet)
        postal_code = self.labWb.getFieldValue('Facility address (postcode)',self.templateWb.activeSheet)
        city = self.labWb.getFieldValue('Facility address (city)',self.templateWb.activeSheet)
        country = self.labWb.getFieldValue('Facility address (country)',self.templateWb.activeSheet)
        address = {'street_with_number' : facility_address,
                   'postal_code' : postal_code,
                   'city' : city,
                   'country': country 
                   }
        
        # gps
        gps_lat = self.labWb.getFieldValue('Facility gpsLat (decimal degree)',self.templateWb.activeSheet)
        gps_lon = self.labWb.getFieldValue('Facility gpsLon (decimal degree)',self.templateWb.activeSheet)
        gps = {'gps_lat': gps_lat,
               'gps_lon': gps_lon
               }
        
        tna = self.labWb.getFieldValue('Facility participates to TNA call? Add TNA call website if YES',self.templateWb.activeSheet)
        
        # affiliation
        # Currently not further elaborated on.
        # The Excel template does not currently request address and PIC e.g.
        affiliation = self.labWb.getFieldValue('Affiliation of Facility contact person',self.templateWb.activeSheet)
        affiliation = {'legal_name': affiliation,
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
        contact_person = {'first_name' : self.labWb.getFieldValue('Facility contact person (first name)',self.templateWb.activeSheet),
                          'family_name': self.labWb.getFieldValue('Facility contact person (family name)',self.templateWb.activeSheet),
                          'identifier': {
                                  'id_type': '',
                                  'id_value': self.labWb.getFieldValue('Facility contact person ID ',self.templateWb.activeSheet)},
                          'email': self.labWb.getFieldValue('Email of Facility contact person',self.templateWb.activeSheet),
                          'affiliation': affiliation}
        
        # facility
        # we use representation names from V6.1 onwards:
        labName = self.labWb.getFieldValue('Facility Name (if other)',self.templateWb.activeSheet)
        
        #***************************************
        # Procedure for choosing identifiers:
        #
        # 1) check on FacilityName or FacilityNameOther in IDsFile:inputstring
        # 2) if not found then generateID with inputstring = labName + ('Affiliation of Facility contact person', City)
        # 3) append new id with inputstring, labname, id to IDsFile
        #
        # generate check on identifiersexport CKAN, connect them to IDsFile for retrieval of original oinputstrings
        # second check with generation of identifiers and compare to #ids in CKAN
        #
        #***************************************

        labNamesForIdRetrieval = []
        labNamesForIdRetrieval.append(self.labWb.getFieldValue('Facility Name',self.templateWb.activeSheet))
        labNamesForIdRetrieval.append(self.labWb.getFieldValue('Facility Name (if other)',self.templateWb.activeSheet))

        
        self.labID = self.labWb.getFieldValue('Facility ID',self.templateWb.activeSheet)
        if self.labID in ['will be assigned later', '']:
            # if not already provided
            self.labID = self.IDs.readLabIdentifier(labNamesForIdRetrieval)
        
        # links to the MSL CKAN catalogue portal
        # We use the TCS Portal Webservice to check whether this lab has data publications already:
        
        numOfPublications = self.ckan.retrieveNumberOfLabPublications(self.labID)
        if numOfPublications > 0:
            dataServices = [{'service_type': 'data_publications_access',
                             'link_label' : 'Go to data publications from this lab (MSL TCS catalogue portal)',
                             'URL': 'https://epos-msl.uu.nl/organization/' + self.labID},
                          {'service_type': 'data_publications_get',
                           'link_label': 'Retrieve data publications from this lab',
                           'URL': 'https://epos-msl.uu.nl/ics/api.php?Lab=' + self.labID,
                           'payload' : 'json'}]
        else:
            dataServices = []
        
        dataServices.append({'service_type': 'TCS_portal_redirection',
                        'link_label': 'More facility information',
                        'URL': 'https://epos-msl.uu.nl/organization/about/' + self.labID})
                
        if tna != '':
            dataServices.append({'service_type': 'TNA_redirection',
                                 'link_label': 'More information about TNA possibilities at this lab',
                                 'URL': tna})
    
    
        if self.labID == '':
            log = {'Missing identifier for' : fileName + ' (labName = ' + labNamesForIdRetrieval[0] + ')'}
            logIdentifiers.append(log)
        else:
            if not self.ckan.identifierInPortal(self.labID):
                log = {'Identifier ' + self.labID + ' for' : fileName + ' (labName = ' + labNamesForIdRetrieval[0] + ') not in portal'}
                logIdentifiers.append(log)
            
        riName = self.labWb.getFieldValue('RI name',self.templateWb.activeSheet)
    
        facility = {'facility' : {'type' : 'laboratory',
                    'lab_id' : self.labID,
                    'research_infrastructure_name' : riName,
                    'facility_name' : labName,
                    'address' : address,
                    'gps' : gps,
                    'website' : self.labWb.getFieldValue('Facility website',self.templateWb.activeSheet),
                    'lab_services' : self.fillLabServices(),
                    'data_services' : dataServices},
                    }
            #'contact_person' : contact_person,
            # 'general_description' : getFieldValue(sheet,'Lab information',templateSheet),
        # TODO generate test identifiers file 
        genIDs = [self.IDs.generateLabIdentifier(labNamesForIdRetrieval[0]),self.IDs.generateLabIdentifier(labNamesForIdRetrieval[0])]
        if self.labID in genIDs:
            self.IDs.matchGenIDs['generated_ids'].append({'name':labNamesForIdRetrieval[0], 'id': genIDs})
        else:
            self.IDs.unmatchGenIDs['generated_ids'].append({'name':labNamesForIdRetrieval[0], 'id': genIDs})
        return facility

    def fillLabServices(self):
        # equipment types
        equipmentTypes = []
        # find header row
        equipmentTypeHeaderRowNum = self.labWb.getCellNumberByStringValue('Equipment type')[0]
        valueRow = equipmentTypeHeaderRowNum+1
        equipmentTemplateTypeHeaderRowNum = self.templateWb.getCellNumberByStringValue('Equipment type')[0]
        valueTemplateRow = equipmentTemplateTypeHeaderRowNum+1
        
        # in the template valueRow contains possible default values against which we must check:
        defaults = []
        for x in range(0, 11):
            defaults.append(self.templateWb.getFieldValueByCellNum((valueTemplateRow,x),[]))
        
        # get separator by checking on merged cells values
        mergedCells = self.labWb.getListOfMergedCellValues()
        mergedCells.append('') # otherwise we'll fail on the empty row
    
        eqType = self.labWb.getFieldValueByCellNum((valueRow,0),defaults)
    
        while (not eqType in mergedCells) and (valueRow < self.sheet.nrows):
            if not self.labWb.checkOnEmptyRow(valueRow):    
                equipment_name = self.labWb.getFieldValueByCellNum((valueRow,2),defaults)
                equipment_custom_name = self.labWb.getFieldValueByCellNum((valueRow,3),defaults)
                if equipment_custom_name != '':
                    # turn base name into group, assign 'other value' to name
                    entrancePart1 = {'equipment_type' : eqType,
                                     'equipment_group' : self.labWb.getFieldValueByCellNum((valueRow,1),defaults),
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
                                     'equipment_group' : self.labWb.getFieldValueByCellNum((valueRow,1),defaults),
                                     'equipment_name' : equipment_name}
                if eqType == '':
                    log = {'empty equipment type in non-empty entrance' : fileName}
                    logInfo.append(log)
                entrancePart2 = {'equipment_brand' : self.labWb.getFieldValueByCellNum((valueRow,4),defaults),
                            'equipment_website' : self.labWb.getFieldValueByCellNum((valueRow,7),defaults),
                            'equipment_specifics_and_comments' : self.labWb.getFieldValueByCellNum((valueRow,8),defaults),
                            'equipment_quantity' : self.labWb.getFieldValueByCellNum((valueRow,9),defaults),
                            'references' : self.labWb.getFieldValueByCellNum((valueRow,10),defaults)}
                
                # The next fields are not exported as of 2019-09-13. We have to check against GPRD
                # 'equipment_contact_person' : getFieldValueByCellNum(sheet,(valueRow,5),defaults),
                # 'email_equipment_contact_person' : getFieldValueByCellNum(sheet,(valueRow,6),defaults),
               
                entrance = {}
                entrance.update(entrancePart1)
                entrance.update(entrancePart2)
                equipmentTypes.append(entrance)
                
            valueRow += 1
            eqType = self.labWb.getFieldValueByCellNum((valueRow,0),defaults)
            
        # We'll sort on equipment names: discussed in Prague Sept 2019
        from operator import itemgetter
        newList = sorted(equipmentTypes, key=itemgetter('equipment_name'))
        equipmentTypes = newList
            
        # measurement types
        measurementTypes = []
        # find header row
        # update 2019-09-13: not all domains provide this, so we have to build in an exception
        
        measurementTypeHeaderRowNum = self.labWb.getCellNumberByStringValue('Measurement type')[0]
        
        if measurementTypeHeaderRowNum != -1:
        
            valueRow = measurementTypeHeaderRowNum+1
            measurementTypeTemplateHeaderRowNum = self.labWb.getCellNumberByStringValue('Measurement type')[0]
            templateValueRow = measurementTypeTemplateHeaderRowNum+1
            defaults.clear()
            for x in range(6):
                defaults.append(self.templateWb.getFieldValueByCellNum((templateValueRow,x),['']))
            
            measurementType = self.labWb.getFieldValueByCellNum((valueRow,0),defaults)
        
            while (not measurementType in mergedCells) and (valueRow < self.sheet.nrows):
                measurement_name = self.labWb.getFieldValueByCellNum((valueRow,2),defaults)
                if measurement_name.lower() == 'other':
                    measurement_custom_name = self.labWb.getFieldValueByCellNum((valueRow,3),defaults)
                    if measurement_custom_name != '':
                        measurement_name = measurement_custom_name
                entrance = {'measurement_type' : measurementType,
                            'measurement_group' : self.labWb.getFieldValueByCellNum((valueRow,1),defaults),
                            'measurement_name' : measurement_name,
                            'measured_type_specifics_and_comments' : self.labWb.getFieldValueByCellNum((valueRow,4),defaults),
                            'references' : self.labWb.getFieldValueByCellNum((valueRow,5),defaults)}
                measurementTypes.append(entrance)
                valueRow += 1
                measurementType = self.labWb.getFieldValueByCellNum((valueRow,0),defaults)
                
            # Sorting
            #from operator import itemgetter
            newList = sorted(measurementTypes, key=itemgetter('measurement_name'))
            measurementTypes = newList
            
        else:
            measurementTypes = []
        
        returnValue = {'research_field': self.research_field, 
                          'subdomain': [self.subdomain],
                          'equipment' : equipmentTypes,
                          'measurement' : measurementTypes}
        
        
        return returnValue
    
    def readLabInfo(self):
        self.labInfo = self.fillGeneralLabInfo()
        
    def getValue(self, keyName):
        return self.labInfo['facility'].get(keyName) # TODO: need exception handling

# ****************************************
# END OF CLASS DEFINITIONS               *
# ****************************************


def processExcel(research_field, subdomain, fullLabfilePath, outputDir, templateFile, IDs):
    fileName = fullLabfilePath.rsplit('/',1)[1]
    #to be sure that we have no space in the filename:
    fileName = fileName.replace(' ', '_')
    book = excelWB(fullLabfilePath)
    template = excelWB(templateFile)
    newLab = lab(book, research_field, subdomain, template, IDs)
    if newLab.labInfo['facility'].get('lab_id') != '':
        # we only add labs with a valid identifier to the allLabs export for ICS
        allLabsExport.append(newLab.labInfo)
    else:
        allLabsNotExported.append(newLab.labInfo)
    writeJSONFile(outputDir + '/' + fileName + '.json',newLab.labInfo)

#---------------
    
def getSources(root):
    import glob
    sourceFiles = glob.glob(root + '[A-Z,a-z]*.xlsx')
    for file in sourceFiles:
        if file.find('TEMPLATE') > -1:
            sourceFiles.remove(file)
    return sourceFiles

# running the conversion script with special parameters from the jupyter notebook:

def runConversion(PALEOTEMPLATE = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/Paleomag/Updated_Paleomag_Lab_Description_V6.2/TEMPLATE_Laboratory description_paleomagnetism_V6.2.xlsx',
                  PALEOFILES = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/Paleomag/Updated_Paleomag_Lab_Description_V6.2/',
                  ROCKTEMPLATE = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/Rock physics/Updated_Rock_Physics_Lab_Description_V3.2/TEMPLATE_Laboratory description_rock physics_V3.2.xlsx',
                  ROCKFILES = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/Rock physics/Updated_Rock_Physics_Lab_Description_V3.2/',
                  ANALYTICALTEMPLATE = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/Analytical&Microscopy/Updated_Analytical&Microscopy_Lab_Description_V3.1/TEMPLATE_Laboratory description_analytical labs_V3.1.xlsx',
                  ANALYTICALFILES = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/Analytical&Microscopy/Updated_Analytical&Microscopy_Lab_Description_V3.1/',
                  ANALOGUETEMPLATE = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/Analogue modelling/Updated_Analogue_Lab_Description_V3/TEMPLATE_Laboratory description_analogue modelling_V3.xlsx',
                  ANALOGUEFILES = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/Analogue modelling/Updated_Analogue_Lab_Description_V3/',
                  IDENTIFIERSFILE = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/lab_identifiers.json',
                  JSONOUTPUT = '/Users/otto/GitLab/inframsl/labs.json',
                  JSONOUTPUT_ISSUES = '/Users/otto/GitLab/inframsl/labs_issues.json'):
    
    domains = [{'research_field': 'Paleomagnetism',
                'subdomain': 'Paleomagnetic and magnetic data',
                'sourceDir' : PALEOFILES,
                'sourceFiles': getSources(PALEOFILES),
                'template': PALEOTEMPLATE},
        {'research_field': 'Rock physics',
                'subdomain': 'Rock and melt physical properties',
                'sourceDir' : ROCKFILES,
                'sourceFiles': getSources(ROCKFILES),
                'template': PALEOTEMPLATE},
         {'research_field': 'Analytical and microscopy',
                'subdomain': 'Analytical and microscopy data',
                'sourceDir' : ANALYTICALFILES,
                'sourceFiles': getSources(ANALYTICALFILES),
                'template': ANALYTICALTEMPLATE},
          {'research_field': 'Analogue modelling',
                'subdomain': 'Analogue modelling of geologic processes',
                'sourceDir' : ANALOGUEFILES,
                'sourceFiles': getSources(ANALOGUEFILES),
                'template': ANALOGUETEMPLATE}]

    
    # we create an object for manipulating lab IDs:
    allIDs = labIDs(IDENTIFIERSFILE)
    
    
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
    
    global allLabsNotExported
    allLabsNotExported = []
    
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
            processExcel(research_field, subdomain,file, outputDir,templateFile,allIDs)
            
    log = []
    log.append(logIdentifiers)
    log.append(logInfo)
    writeJSONFile('./log.json',log)
    infraStructures = {'infrastructures' : allLabsExport}
    # writeJSONFile(outputDir + '/allpaleomagLabs_v6.1.json', infraStructures)
    # and write the new source for the service:
    writeJSONFile(JSONOUTPUT, infraStructures)
    notExported = {'infrastructures' : allLabsNotExported}
    writeJSONFile(JSONOUTPUT_ISSUES, notExported)




#---------------
    
# running the whole script with standard parameters:
    
if __name__ == "__main__":
    
    PALEOTEMPLATE = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/Paleomag/Updated_Paleomag_Lab_Description_V6.2/TEMPLATE_Laboratory description_paleomagnetism_V6.2.xlsx'
    PALEOFILES = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/Paleomag/Updated_Paleomag_Lab_Description_V6.2/'
    
    ROCKTEMPLATE = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/Rock physics/Updated_Rock_Physics_Lab_Description_V3.2/TEMPLATE_Laboratory description_rock physics_V3.2.xlsx'
    ROCKFILES = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/Rock physics/Updated_Rock_Physics_Lab_Description_V3.2/'
    
    ANALYTICALTEMPLATE = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/Analytical&Microscopy/Updated_Analytical&Microscopy_Lab_Description_V3.1/TEMPLATE_Laboratory description_analytical labs_V3.1.xlsx'
    ANALYTICALFILES = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/Analytical&Microscopy/Updated_Analytical&Microscopy_Lab_Description_V3.1/'
    
    ANALOGUETEMPLATE = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/Analogue modelling/Updated_Analogue_Lab_Description_V3/TEMPLATE_Laboratory description_analogue modelling_V3.xlsx'
    ANALOGUEFILES = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/Analogue modelling/Updated_Analogue_Lab_Description_V3/'
    
    IDENTIFIERSFILE = '/Users/otto/ownCloud - EPOS/WP16/LABS description service/Lab info collected/lab_identifiers.json'

    # we create an object for manipulating lab IDs:
    allIDs = labIDs(IDENTIFIERSFILE)
    
    domains = [{'research_field': 'Paleomagnetism',
                'subdomain': 'Paleomagnetic and magnetic data',
                'sourceDir' : PALEOFILES,
                'sourceFiles': getSources(PALEOFILES),
                'template': PALEOTEMPLATE},
        {'research_field': 'Rock physics',
                'subdomain': 'Rock and melt physical properties',
                'sourceDir' : ROCKFILES,
                'sourceFiles': getSources(ROCKFILES),
                'template': PALEOTEMPLATE},
         {'research_field': 'Analytical and microscopy',
                'subdomain': 'Analytical and microscopy data',
                'sourceDir' : ANALYTICALFILES,
                'sourceFiles': getSources(ANALYTICALFILES),
                'template': ANALYTICALTEMPLATE},
          {'research_field': 'Analogue modelling',
                'subdomain': 'Analogue modelling of geologic processes',
                'sourceDir' : ANALOGUEFILES,
                'sourceFiles': getSources(ANALOGUEFILES),
                'template': ANALOGUETEMPLATE}]
    
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
    
    global allLabsNotExported
    allLabsNotExported = []
    
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
            processExcel(research_field, subdomain,file, outputDir,templateFile,allIDs)
            
    log = []
    log.append(logIdentifiers)
    log.append(logInfo)
    writeJSONFile('./log.json',log)
    infraStructures = {'infrastructures' : allLabsExport}
    # writeJSONFile(outputDir + '/allpaleomagLabs_v6.1.json', infraStructures)
    # and write the new source for the service:
    writeJSONFile('/Users/otto/GitLab/inframsl/labs.json', infraStructures)
    notExported = {'infrastructures' : allLabsNotExported}
    writeJSONFile('/Users/otto/GitLab/inframsl/labs_issues.json', notExported)
    writeJSONFile('/Users/otto/GitLab/inframsl/ids_generated_match.json', allIDs.matchGenIDs)
    writeJSONFile('/Users/otto/GitLab/inframsl/ids_generated_unmatch.json', allIDs.unmatchGenIDs)
