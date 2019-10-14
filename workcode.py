#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Oct 14 11:49:44 2019

@author: otto
"""

if __name__ == "__main__":

    import custom_json
    
    idFile = 'lab_identifiers.json'
    
    myList = custom_json.loadJSONFromFile(idFile)
    
    newList = []
    
    for entry in myList:
        newItem = {}
        newItem.update({"domain":entry['domain']})
        newItem.update({"inputstring":entry['name']})
        newItem.update({"labname":entry['name']})
        print(entry['name'])
        newItem.update({"affiliation":entry['affiliation']})
        newItem.update({"id":entry['id']})
        newList.append(newItem)
        
    custom_json.writeJSONFile('newList.json',newList)
    
    