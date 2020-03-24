#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Sep 14 17:13:03 2019

@author: otto
"""

import urllib
import urllib.parse
import pprint

from lib.json_io import *
import requests
import untangle


class TCS_PortalRequests:
    def __init__(self, TCS_PORTAL_API_BASE='https://epos-msl.uu.nl/ics/api.php?',
                 TCS_PORTAL_CKAN_API_BASE='https://epos-msl.uu.nl/api/3/action/'):
        self.TCS_PORTAL_API_BASE = TCS_PORTAL_API_BASE  # custom API
        self.TCS_PORTAL_CKAN_API_BASE = TCS_PORTAL_CKAN_API_BASE  # tag_list'
        self.idsInPortal = self.retrieveAllIdentifiers()  # list
        self.APIkey = '612f3971-3794-4548-9882-aea17dbbe9e9'  # otto@epos-msl.uu.nl

    def retrieveNumberOfLabPublications(self, labId):
        portalRequest = requests.get(self.TCS_PORTAL_API_BASE + 'Lab=' + labId)
        portalRequest.encoding = 'UTF-8'
        json_payload = json.loads(portalRequest.text)
        number_of_data_publications = json_payload['result']['count']
        print('Data publications: ' + str(number_of_data_publications))
        return number_of_data_publications

    def retrieveAllIdentifiers(self):
        url = self.TCS_PORTAL_CKAN_API_BASE + 'organization_list'
        portalRequest = requests.get(url)
        portalRequest.encoding = 'UTF-8'
        json_payload = json.loads(portalRequest.text)
        identifiersAsList = json_payload['result']
        return identifiersAsList

    def getLabDescription(self, labId):
        # http://epos-msl.uu.nl/api/3/action/organization_show?id=9ba34c109b827b177aab36e0266b1643
        portalRequest = requests.get(self.TCS_PORTAL_CKAN_API_BASE + 'organization_show?id=' + labId)
        portalRequest.encoding = 'UTF-8'
        json_payload = json.loads(portalRequest.text)
        json_out = get_json_pretty_dump_as_string(json_payload)
        # we have to convert the unicode escape characters from the json payload to readable characters:
        # encode to make bytes, use codecs to decode again
        json_out.encode('unicode_escape')
        import codecs
        json_out = codecs.decode(json_out, 'unicode_escape')
        return json_out

    def getJSONLabDescription(self, labId):
        # http://epos-msl.uu.nl/api/3/action/organization_show?id=9ba34c109b827b177aab36e0266b1643
        portalRequest = requests.get(self.TCS_PORTAL_CKAN_API_BASE + 'organization_show?id=' + labId)
        portalRequest.encoding = 'UTF-8'
        json_payload = json.loads(portalRequest.text)
        return json_payload

    def getPublicationsRecord(self, labId):
        portalRequest = requests.get(self.TCS_PORTAL_API_BASE + 'Lab=' + labId)
        portalRequest.encoding = 'UTF-8'
        json_payload = json.loads(portalRequest.text)
        json_out = get_json_pretty_dump_as_string(json_payload)
        # we have to convert the unicode escape characters from the json payload to readable characters:
        # encode to make bytes, use codecs to decode again
        json_out.encode('unicode_escape')
        import codecs
        readableText = codecs.decode(json_out, 'unicode_escape')
        return readableText

    def getAllKeyWords(self):
        portalRequest = requests.get(self.TCS_PORTAL_CKAN_API_BASE + 'tag_list')
        portalRequest.encoding = 'UTF-8'
        json_payload = json.loads(portalRequest.text)
        json_out = get_json_pretty_dump_as_string(json_payload)
        # we have to convert the unicode escape characters from the json payload to readable characters:
        # encode to make bytes, use codecs to decode again
        json_out.encode('unicode_escape')
        import codecs
        readableText = codecs.decode(json_out, 'unicode_escape')
        return json_payload

    def loadIdentifiersFile(self, IDsFile):
        f = open(IDsFile)
        returnValue = json.load(f)
        f.close()
        return returnValue

    def identifierInPortal(self, id):
        if id in self.idsInPortal:
            return True
        else:
            return False


class CSIC_XML2JSON:
    def __init__(self, xml_file):
        self.xml_object = untangle.parse(xml_file)
        self.csic_dict = {}
        # fields:
        self.title = ''
        self.description = ''
        self.keywords = []
        self.source = ''
        self.author = []
        self.maintainer = 'DIGITAL.CSIC'  # maintainer
        self.cites = []
        self.dataset_contact = []
        self.is_supplement_to = []
        self.is_referenced_by = []
        self.cites = []
        self.publication_date = ''
        self.publisher = 'https://digital.csic.es/'
        self.geobox_eLong = ''  # we had to fix this manually in the CSIC exports
        self.geobox_nLat = ''
        self.geobox_sLat = ''
        self.geobox_wLong = ''
        self.organization = 'other-lab'  # we do not have the lab associations from CSIC yet
        # local CKAN settings:
        self.API_KEY = '65d1279e-5efb-495e-849c-660c48f9516c'
        self.instance = 'http://localhost:5000'
        self.action_base = self.instance + '/api/action/'

    def set_field_values(self):
        # following the standard MSL form in CKAN
        # DOI
        self.source = self.xml_object.resource.identifier.cdata
        # authors
        for creatorName in self.xml_object.resource.creators.creatorName:
            self.author.append(creatorName.cdata)
        # title
        self.title = self.xml_object.resource.titles.title.cdata
        # abstract
        self.description = "\n\n".join(
            [descriptions.description.cdata for descriptions in self.xml_object.resource.descriptions])
        # keywords
        for subject in self.xml_object.resource.subjects.subject:
            self.keywords.append(subject.cdata)
        # coordinates; file must be prepared!!
        self.geobox_eLong = self.xml_object.resource.custom_geo.eLong.cdata
        self.geobox_nLat = self.xml_object.resource.custom_geo.nLat.cdata
        self.geobox_sLat = self.xml_object.resource.custom_geo.sLat.cdata
        self.geobox_wLong = self.xml_object.resource.custom_geo.wLong.cdata
        # contact person
        for contributors in self.xml_object.resource.contributors:
            if contributors.contributor["contributorType"] == 'contactPerson':
                contact = contributors.contributor.cdata
                self.dataset_contact.append(contact)
            elif contributors.contributor["contributorType"] == 'hostingInstitution':
                lab = contributors.contributor.cdata
                self.organization = lab  # we have to add the identifiers later on!
        # publication data
        self.publication_date = self.xml_object.resource.date.cdata # TODO: check on attribute type 'issued'

    def create_json(self):
        my_dict = {
            "title": self.title,
            "description": self.description,
            "tags": self.keywords,
            "Source": self.source,
            "geobox-eLong": self.geobox_eLong,
            "geobox-nLat": self.geobox_nLat,
            "geobox-sLat": self.geobox_sLat,
            "geobox-wLong": self.geobox_wLong,
            "Author": self.author,
            "organization": self.organization,
            "Dataset contact": self.dataset_contact,
            "Provided by": self.maintainer,
            "Is supplement to": self.is_supplement_to,
            "Is referenced by": self.is_referenced_by,
            "Cites": self.cites,
            "Publisher": self.publisher,
            "Publication date": self.publication_date
        }
        self.csic_dict.update(my_dict)

    def write_json(self, output_file):
        self.create_json()
        write_json_file(output_file,self.csic_dict)

    def write_to_ckan(self):
        request = self.action_base + 'package_create'
        portal_request = requests.put(request)
        dataset_dict = {
            'name': 'testset 2',
            'notes': 'A long description of my dataset',
            'owner_org': 'other-lab'
        }

        # Use the json module to dump the dictionary to a string for posting.
        data_string = urllib.parse.quote(json.dumps(dataset_dict))

        # We'll use the package_create function to create a new dataset.
        request = requests.request("http://localhost:5000/api/action/package_create")

        # Creating a dataset requires an authorization header.
        # Replace *** with your API key, from your user account on the CKAN site
        # that you're creating the dataset on.
        request.headers.update({'Authorization': '65d1279e-5efb-495e-849c-660c48f9516c'})

        # Make the HTTP request.
        response = urllib.urlopen(request, data_string)
        assert response.code == 200

        # Use the json module to load CKAN's response into a dictionary.
        response_dict = json.loads(response.read())
        assert response_dict['success'] is True

        # package_create returns the created package as its result.
        created_package = response_dict['result']
        pprint.pprint(created_package)
