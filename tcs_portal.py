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

def retrieveNumberOfLabPublications(labId):
    portalRequest = requests.get(TCS_PORTAL_API_BASE + 'Lab=' + labId)
    portalRequest.encoding = 'UTF-8'
    json_payload = json.loads(portalRequest.text)
    json_out = custom_json.getJSONPrettyDumpAsString(json_payload)
    # we have to convert the unicode escape characters from the json payload to readbale characters:
    # encode to make bytes, use codecs to decode again
    json_out.encode('unicode_escape')
    import codecs
    readableText = codecs.decode(json_out, 'unicode_escape')
    print(readableText)
    number_of_data_publications = json_payload['result']['count']
    print(number_of_data_publications)
    return number_of_data_publications

    