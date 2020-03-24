#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Sep 11 17:26:11 2019

@author: otto
"""
import glob
import os

def executeDir(dirName):
    files = glob.glob(dirName + '*.docx')
    for i in files:
        cmd = 'pandoc -f docx -t markdown_strict ' + i + ' -o ' + i + '.markdown'
        os.system(cmd)