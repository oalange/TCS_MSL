
# *************************************************
# class for managing identifiers for laboratories *
# *************************************************

"""
As of the start of the collaboration with the MagIC repository for paleomagnetic data, the structure of
the json laboratory info file is frozen. There are two files, one for lab-identification and one for lab-description:

lab identifiers: /Users/otto/GitLab/inframsl/labnames.json
structure example:

[
    {
        "domain": "Analogue modelling of geologic processes",
        "id": "1b9abbf97c7caa2d763b647d476b2910",
        "lab_portal_name": "TecMOD - FASTmodel (CNRS-Paris Sud Orsay University, France)",
        "lab_editor_name": "FASTmodel- Laboratoire de modélisation analogique Fluides Automatique et Systèmes Thermiques (CNRS-Paris Sud Orsay University, France)",
        "affiliation": "CNRS-Paris Sud Orsay University, Paris, France",
        "id_inputstring": "FASTmodel- Laboratoire de mod\u00e9lisation analogique\u00a0Fluides Automatique et Syst\u00e8mes Thermiques (CNRS-Paris Sud Orsay University,  France)"
    },

lab descriptions: /Users/otto/GitLab/inframsl/labs.json

"""

import hashlib
from lib.json_io import *

class labIDs:
    def __init__(self, identifiers_file='/Users/otto/GitLab/inframsl/labnames.json'):
        #self.dictIDs = {'allIDs': []}  # all template-based generated IDs
        self.identifiers_file = identifiers_file
        self.labs = load_json_from_file(self.identifiers_file)

    def readLabIdentifier(self, labNames):  # labNames is now a list of multiple possible candidate strings
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
                    name = lab["id_inputstring"].replace('  ', ' ')
                    name = name.replace('"', '')
                    name = name.replace(',', '')
                    name = name.lower()
                    if name.find(labNameStripped) != -1:
                        # when the name provided in the template is a non-empty substring then we have a hit
                        labId = lab["id"]
                        NotFound = False

                if NotFound:  # we do a check on the labname field
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
        entries = load_json_from_file(idFile)
        for entry in entries:
            updatedEntry = entry
            updatedEntry.update({'id': self.generateLabIdentifier(entry['id_inputstring'])})
            output.append(updatedEntry)
        write_json_file('out_' + idFile, output)

    def addEmptyLabIdentifiers(self, new_output_file, input_file=''):
        # expects a valid identifiersFile and fills/updates the missing id-fields
        if input_file == '':
            input_file = self.identifiers_file
        entries = load_json_from_file(input_file)
        for entry in entries:
            if entry['id'] == '':
                new_id = self.generateLabIdentifier(entry['id_inputstring'])
                entry.update({'id': new_id})
        write_json_file(new_output_file, entries)

    def checkIdentifiers(self):
        # checks identifiers against inputstrings
        sourceList = load_json_from_file(self.identifiers_file)
        for lab in sourceList:
            check_id = self.generateLabIdentifier(lab['id_inputstring'])
            if check_id != lab['id']:
                print('Labname ' + lab['lab_editor_name'] + ' has id mismatch')
