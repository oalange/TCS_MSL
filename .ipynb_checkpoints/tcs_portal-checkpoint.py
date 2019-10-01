#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Sep 14 17:13:03 2019

@author: otto
"""

import json
import custom_json
import requests

TCS_PORTAL_API_BASE = 'https://epos-msl.uu.nl/ics/api.php?'
TCS_PORTAL_CKAN_API_BASE = 'https://epos-msl.uu.nl/api/3/action/' #tag_list'

def retrieveNumberOfLabPublications(labId):
    portalRequest = requests.get(TCS_PORTAL_API_BASE + 'Lab=' + labId)
    portalRequest.encoding = 'UTF-8'
    json_payload = json.loads(portalRequest.text)
    number_of_data_publications = json_payload['result']['count']
    print('Data publications: ' + str(number_of_data_publications))
    return number_of_data_publications

def retrieveAllIdentifiers():
    url = TCS_PORTAL_CKAN_API_BASE + 'organization_list'
    portalRequest = requests.get(url)
    portalRequest.encoding = 'UTF-8'
    json_payload = json.loads(portalRequest.text)
    identifiersAsList = json_payload['result']
    return identifiersAsList

def getLabDescription(labId):
    #http://epos-msl.uu.nl/api/3/action/organization_show?id=9ba34c109b827b177aab36e0266b1643
    portalRequest = requests.get(TCS_PORTAL_CKAN_API_BASE + 'organization_show?id=' + labId)
    portalRequest.encoding = 'UTF-8'
    json_payload = json.loads(portalRequest.text)
    json_out = custom_json.getJSONPrettyDumpAsString(json_payload)
    # we have to convert the unicode escape characters from the json payload to readable characters:
    # encode to make bytes, use codecs to decode again
    json_out.encode('unicode_escape')
    import codecs
    json_out = codecs.decode(json_out, 'unicode_escape')
    return json_out

def getJSONLabDescription(labId):
    #http://epos-msl.uu.nl/api/3/action/organization_show?id=9ba34c109b827b177aab36e0266b1643
    portalRequest = requests.get(TCS_PORTAL_CKAN_API_BASE + 'organization_show?id=' + labId)
    portalRequest.encoding = 'UTF-8'
    json_payload = json.loads(portalRequest.text)
    return json_payload

def getPublicationsRecord(labId):
    portalRequest = requests.get(TCS_PORTAL_API_BASE + 'Lab=' + labId)
    portalRequest.encoding = 'UTF-8'
    json_payload = json.loads(portalRequest.text)
    json_out = custom_json.getJSONPrettyDumpAsString(json_payload)
    # we have to convert the unicode escape characters from the json payload to readable characters:
    # encode to make bytes, use codecs to decode again
    json_out.encode('unicode_escape')
    import codecs
    readableText = codecs.decode(json_out, 'unicode_escape')
    return readableText

def getAllKeyWords():
    portalRequest = requests.get(TCS_PORTAL_CKAN_API_BASE + 'tag_list')
    portalRequest.encoding = 'UTF-8'
    json_payload = json.loads(portalRequest.text)
    json_out = custom_json.getJSONPrettyDumpAsString(json_payload)
    # we have to convert the unicode escape characters from the json payload to readable characters:
    # encode to make bytes, use codecs to decode again
    json_out.encode('unicode_escape')
    import codecs
    readableText = codecs.decode(json_out, 'unicode_escape')
    return json_payload

def loadIdentifiersFile(IDsFile):
    f = open(IDsFile)
    returnValue = json.load(f)
    f.close()
    return returnValue

    