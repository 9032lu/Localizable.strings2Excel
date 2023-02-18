#!/usr/bin/python
# -*- coding: UTF-8 -*-

import openpyxl
# import sys
import os
# import imaplib, sys

class XlsFileUtil:
    'xls file util'

    def __init__(self, filePath):
        self.filePath = filePath
        # get all sheets
        # imaplib.reload(sys)
        # reload(sys)
        # sys.setdefaultencodin
        # g('utf-8')
        self.workBook = openpyxl.load_workbook(filePath)

    def getAllTables(self):
        return self.workBook.worksheet

    def getTableByIndex(self, index):
        if index >= 0 and index < len(self.workBook.worksheets):
            return self.workBook.worksheets[index]
        else:
            print("XlsFileUtil error -- getTable:index")

    def getTableByName(self, name):
        return self.workBook.get_sheet_by_name(name)
