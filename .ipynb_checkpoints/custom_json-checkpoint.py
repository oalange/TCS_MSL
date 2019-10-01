#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Sep 14 17:13:52 2019

@author: otto
"""

try:
    import simplejson as json
except ImportError:
    import json

# managing JSON I/O
    
def getJSONPrettyDumpAsString(myDict):
    # returns a pretty formatted JSON fragment from a Python dictionary
    jsonDump = json.dumps(myDict, indent=4, separators=(', ', ': '))
    jsonDump.encode('unicode_escape')
    import codecs
    jsonDump = codecs.decode(jsonDump, 'unicode_escape')
    return jsonDump
    
def loadJSONFromFile(fileName):
    # returns JSON-file's content as Python dictionary
    with open(fileName) as f:
        jsonDict = json.load(f)
        return jsonDict

def writeJSONFile(fileName, jsonData):
    with open(fileName, 'w', encoding='utf-8') as f:
        json.dump(jsonData, f, ensure_ascii=False, indent=4)