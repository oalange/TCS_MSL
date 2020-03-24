from .lab_excel import excelWB as excel_wb
from .lab_id import labIDs as labIdentifier
from infra_msl.msl_portal.tcs_portal import TCS_PortalRequests as portal_msl
from lib.json_io import *
import logging

class labDescription:

    def __init__(self, lab_excel_wb, research_field, subdomain, template_excel_wb, ids_file):
        self.msl_catalogue = portal_msl()
        self.current_excel_file_name = lab_excel_wb
        self.lab_excel_wb = excel_wb(lab_excel_wb)
        self.excel_sheet = self.lab_excel_wb.activeSheet
        self.research_field = research_field
        self.subdomain = subdomain
        self.template_excel_wb = excel_wb(template_excel_wb)
        self.ids = labIdentifier(ids_file)
        self.lab_name = ''
        self.ri_name = ''
        self.address = dict()
        self.affiliation = dict()
        self.contact_person = dict()
        self.gps = dict()
        self.tna = ''
        self.lab_website = ''
        self.data_services = list()
        self.lab_services = dict()
        self.facility_info = dict()
        self.labID = ''
        logging.basicConfig(filename='lab_info.log', level=logging.INFO)
        self.update_facility_info()

    def fillGeneralLabInfo(self):

        # ********************
        # address
        # ********************

        facility_address = self.lab_excel_wb.getFieldValue('Facility address (street + number)', self.template_excel_wb.activeSheet)
        postal_code = self.lab_excel_wb.getFieldValue('Facility address (postcode)', self.template_excel_wb.activeSheet)
        city = self.lab_excel_wb.getFieldValue('Facility address (city)', self.template_excel_wb.activeSheet)
        country = self.lab_excel_wb.getFieldValue('Facility address (country)', self.template_excel_wb.activeSheet)
        self.address = {'street_with_number': facility_address,
                   'postal_code': postal_code,
                   'city': city,
                   'country': country
                   }

        # ********************
        # gps
        # ********************

        gps_lat = self.lab_excel_wb.getFieldValue('Facility gpsLat (decimal degree)', self.template_excel_wb.activeSheet)
        gps_lon = self.lab_excel_wb.getFieldValue('Facility gpsLon (decimal degree)', self.template_excel_wb.activeSheet)
        self.gps = {'gps_lat': gps_lat,
               'gps_lon': gps_lon
               }

        # ********************************
        # affiliation
        # ********************************

        # Currently not further elaborated on.
        # The Excel template does not currently request address and PIC e.g.
        affiliation = self.lab_excel_wb.getFieldValue('Affiliation of Facility contact person', self.template_excel_wb.activeSheet)
        self.affiliation = {'legal_name': affiliation,
                       'identifier': {
                           'id_type': '',
                           'id_value': ''},
                       'address': {
                           'street_with_number': '',
                           'postal_code': '',
                           'city': '',
                           'country': ''}
                       }

        # ********************************
        # contact person
        # ********************************

        # not used by now because of GDPR
        """
        contact_person = {
            'first_name': self.labWb.getFieldValue('Facility contact person (first name)', self.templateWb.activeSheet),
            'family_name': self.labWb.getFieldValue('Facility contact person (family name)',
                                                    self.templateWb.activeSheet),
            'identifier': {
                'id_type': '',
                'id_value': self.labWb.getFieldValue('Facility contact person ID ', self.templateWb.activeSheet)},
            'email': self.labWb.getFieldValue('Email of Facility contact person', self.templateWb.activeSheet),
            'affiliation': affiliation}
        """

        # ********************************
        # facility naming
        # ********************************

        # we use representation names from V6.1 onwards:
        self.lab_name = self.lab_excel_wb.getFieldValue('Facility Name (if other)', self.template_excel_wb.activeSheet)
        self.ri_name = self.lab_excel_wb.getFieldValue('RI name', self.template_excel_wb.activeSheet)
        self.lab_website = self.lab_excel_wb.getFieldValue('Facility website', self.template_excel_wb.activeSheet)

        # ********************************
        # lab services
        # ********************************

        self.get_lab_services()

        # ********************************
        # lab identifier
        # ********************************

        self.get_identifier()

        # ********************************
        # data services
        # ********************************

        self.get_data_services()

        # ********************************
        # tna services
        # ********************************

        self.get_tna_services()

        # ********************************
        # putting it all together: facility entry
        # ********************************

        self.facility_info = {'facility': {'type': 'laboratory',
                                 'lab_id': self.labID,
                                 'research_infrastructure_name': self.ri_name,
                                 'facility_name': self.lab_name,
                                 'address': self.address,
                                 'gps': self.gps,
                                 'website': self.lab_website,
                                 'lab_services': self.lab_services,
                                 'data_services': self.data_services},
                    }
        return 0

    # ***********************************************************************************************
    """
        Procedure:
        1) check on FacilityName or FacilityNameOther by comparing with IDsFile
        2) if not found then generateID with inputstring = lab_name + ('Affiliation of Facility contact person', City)
        3) append new id with id_inputstring, labname, id to IDsFile

        TODO: generate check on identifiers export CKAN, connect them to IDsFile for retrieval of original inputstrings
        second check with generation of identifiers and compare to #ids in CKAN
    """

    def get_identifier(self):
        labNamesForIdRetrieval = list() # because we can have more than one option from the sheet we use a list
        labNamesForIdRetrieval.append(self.lab_excel_wb.getFieldValue('Facility Name', self.template_excel_wb.activeSheet))
        labNamesForIdRetrieval.append(self.lab_excel_wb.getFieldValue('Facility Name (if other)', self.template_excel_wb.activeSheet))

        self.labID = self.lab_excel_wb.getFieldValue('Facility ID', self.template_excel_wb.activeSheet)
        if self.labID in ['will be assigned later', '']:
            # thus if not already provided through sheet
            self.labID = self.ids.readLabIdentifier(labNamesForIdRetrieval)

        if self.labID == '':
            log = {'Missing identifier for': self.current_excel_file_name + ' (lab_name = ' + labNamesForIdRetrieval[0] + ')'}
            logging.info(log)
        else:
            if not self.msl_catalogue.identifierInPortal(self.labID):
                log = {'Identifier ' + self.labID + ' for': self.current_excel_file_name + ' (lab_name = ' + labNamesForIdRetrieval[
                    0] + ') not in portal'}
                logging.info(log)

    # ***********************************************************************************************

    def get_tna_services(self):
        self.tna = self.lab_excel_wb.getFieldValue('Facility participates to TNA call? Add TNA call website if YES',
                                                   self.template_excel_wb.activeSheet)
        if self.tna != '':
            self.data_services.append({'service_type': 'TNA_redirection',
                                 'link_label': 'More information about TNA possibilities at this lab',
                                 'URL': self.tna})

    # ***********************************************************************************************

    def get_data_services(self):
        # TODO: check on validity of self.labID before connecting to the portal
        # links to the MSL CKAN catalogue portal
        # We use the TCS Portal Webservice to check whether this lab has data publications already:

        numOfPublications = self.msl_catalogue.retrieveNumberOfLabPublications(self.labID)
        if numOfPublications > 0:
            self.data_services = [{'service_type': 'data_publications_access',
                             'link_label': 'Go to data publications from this lab (MSL TCS catalogue portal)',
                             'URL': 'https://epos-msl.uu.nl/organization/' + self.labID},
                                  {'service_type': 'data_publications_get',
                             'link_label': 'Retrieve data publications from this lab',
                             'URL': 'https://epos-msl.uu.nl/ics/api.php?Lab=' + self.labID,
                             'payload': 'json'}]

        self.data_services.append({'service_type': 'TCS_portal_redirection',
                             'link_label': 'More facility information',
                             'URL': 'https://epos-msl.uu.nl/organization/about/' + self.labID})

    # ***********************************************************************************************

    def get_lab_services(self):
        # equipment types
        equipmentTypes = list()
        # find header row
        equipmentTypeHeaderRowNum = self.lab_excel_wb.getCellNumberByStringValue('Equipment type')[0]
        valueRow = equipmentTypeHeaderRowNum + 1
        equipmentTemplateTypeHeaderRowNum = self.template_excel_wb.getCellNumberByStringValue('Equipment type')[0]
        valueTemplateRow = equipmentTemplateTypeHeaderRowNum + 1

        # in the template valueRow contains possible default values against which we must check:
        defaults = list()
        for x in range(0, 11):
            defaults.append(self.template_excel_wb.getFieldValueByCellNum((valueTemplateRow, x), []))

        # get separator by checking on merged cells values
        mergedCells = self.lab_excel_wb.getListOfMergedCellValues()
        mergedCells.append('')  # otherwise we'll fail on the empty row

        eqType = self.lab_excel_wb.getFieldValueByCellNum((valueRow, 0), defaults)

        while (not eqType in mergedCells) and (valueRow < self.excel_sheet.nrows):
            if not self.lab_excel_wb.checkOnEmptyRow(valueRow):
                equipment_name = self.lab_excel_wb.getFieldValueByCellNum((valueRow, 2), defaults)
                equipment_custom_name = self.lab_excel_wb.getFieldValueByCellNum((valueRow, 3), defaults)
                if equipment_custom_name != '':
                    # turn base name into group, assign 'other value' to name
                    entrancePart1 = {'equipment_type': eqType,
                                     'equipment_group': self.lab_excel_wb.getFieldValueByCellNum((valueRow, 1), defaults),
                                     'equipment_name': equipment_name,
                                     'equipment_secondary_name': equipment_custom_name}
                    log = {'domain': 'paleomag',
                           'category': 'equipment',
                           'File': self.current_excel_file_name,
                           'base_name': equipment_name,
                           'specific name': equipment_custom_name}
                    logging.info(log)

                else:
                    entrancePart1 = {'equipment_type': eqType,
                                     'equipment_group': self.lab_excel_wb.getFieldValueByCellNum((valueRow, 1), defaults),
                                     'equipment_name': equipment_name}
                if eqType == '':
                    log = {'empty equipment type in non-empty entrance': self.current_excel_file_name}
                    logging.error(log)
                entrancePart2 = {'equipment_brand': self.lab_excel_wb.getFieldValueByCellNum((valueRow, 4), defaults),
                                 'equipment_website': self.lab_excel_wb.getFieldValueByCellNum((valueRow, 7), defaults),
                                 'equipment_specifics_and_comments': self.lab_excel_wb.getFieldValueByCellNum((valueRow, 8),
                                                                                                              defaults),
                                 'equipment_quantity': self.lab_excel_wb.getFieldValueByCellNum((valueRow, 9), defaults),
                                 'references': self.lab_excel_wb.getFieldValueByCellNum((valueRow, 10), defaults)}

                # The next fields are not exported as of 2019-09-13. We have to check against GPRD
                # 'equipment_contact_person' : getFieldValueByCellNum(sheet,(valueRow,5),defaults),
                # 'email_equipment_contact_person' : getFieldValueByCellNum(sheet,(valueRow,6),defaults),

                entrance = {}
                entrance.update(entrancePart1)
                entrance.update(entrancePart2)
                equipmentTypes.append(entrance)

            valueRow += 1
            eqType = self.lab_excel_wb.getFieldValueByCellNum((valueRow, 0), defaults)

        # We'll sort on equipment names: discussed in Prague Sept 2019
        from operator import itemgetter
        newList = sorted(equipmentTypes, key=itemgetter('equipment_name'))
        equipmentTypes = newList

        # measurement types
        measurementTypes = list()
        # find header row
        # update 2019-09-13: not all domains provide this, so we have to build in an exception

        measurementTypeHeaderRowNum = self.lab_excel_wb.getCellNumberByStringValue('Measurement type')[0]

        if measurementTypeHeaderRowNum != -1:

            valueRow = measurementTypeHeaderRowNum + 1
            measurementTypeTemplateHeaderRowNum = self.lab_excel_wb.getCellNumberByStringValue('Measurement type')[0]
            templateValueRow = measurementTypeTemplateHeaderRowNum + 1
            defaults.clear()
            for x in range(6):
                defaults.append(self.template_excel_wb.getFieldValueByCellNum((templateValueRow, x), ['']))

            measurementType = self.lab_excel_wb.getFieldValueByCellNum((valueRow, 0), defaults)

            while (not measurementType in mergedCells) and (valueRow < self.excel_sheet.nrows):
                measurement_name = self.lab_excel_wb.getFieldValueByCellNum((valueRow, 2), defaults)
                if measurement_name.lower() == 'other':
                    measurement_custom_name = self.lab_excel_wb.getFieldValueByCellNum((valueRow, 3), defaults)
                    if measurement_custom_name != '':
                        measurement_name = measurement_custom_name
                entrance = {'measurement_type': measurementType,
                            'measurement_group': self.lab_excel_wb.getFieldValueByCellNum((valueRow, 1), defaults),
                            'measurement_name': measurement_name,
                            'measured_type_specifics_and_comments': self.lab_excel_wb.getFieldValueByCellNum((valueRow, 4),
                                                                                                             defaults),
                            'references': self.lab_excel_wb.getFieldValueByCellNum((valueRow, 5), defaults)}
                measurementTypes.append(entrance)
                valueRow += 1
                measurementType = self.lab_excel_wb.getFieldValueByCellNum((valueRow, 0), defaults)

            # Sorting
            # from operator import itemgetter
            newList = sorted(measurementTypes, key=itemgetter('measurement_name'))
            measurementTypes = newList

        else:
            measurementTypes = []

        self.lab_services = {'research_field': self.research_field,
                       'subdomain': [self.subdomain],
                       'equipment': equipmentTypes,
                       'measurement': measurementTypes}


    # ***********************************************************************************************

    def update_facility_info(self):
        self.fillGeneralLabInfo()

    # ***********************************************************************************************

    def get_facility_field_value(self, keyName):
        return self.facility_info['facility'].get(keyName)  # TODO: needs exception handling

    def facility_info_to_file(self, output_file):
        write_json_file(output_file, self.facility_info)

    """
    def check_identifiers(self):
        # TODO generate test identifiers file
        genIDs = [self.ids_file.generateLabIdentifier(labNamesForIdRetrieval[0]),
                  self.ids_file.generateLabIdentifier(labNamesForIdRetrieval[0])]
        if self.labID in genIDs:
            self.ids_file.matchGenIDs['generated_ids'].append({'name': labNamesForIdRetrieval[0], 'id': genIDs})
        else:
            self.ids_file.unmatchGenIDs['generated_ids'].append({'name': labNamesForIdRetrieval[0], 'id': genIDs})
    """

def create_info_from_json(json_file):
    # json-file is a dictionary of parameters, e.g. the example.json in the package dir
    parameters = load_json_from_file(json_file)
    print(parameters)
    try:
        lab_excel_wb = parameters['lab_excel_wb']
        research_field = parameters['research_field']
        subdomain = parameters['subdomain']
        template_excel_wb = parameters['template_excel_wb']
        ids_file = parameters['ids_file']
        new_lab = labDescription(lab_excel_wb, research_field, subdomain, template_excel_wb, ids_file)
        return new_lab
    except KeyError:
        print("KeyError in json")
    except:
        print("Something else went wrong")

