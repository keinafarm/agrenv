# -*- coding: utf-8 -*-
from Debug import Debug
############################################################
#
#   1行分のExcelデータを保持する
#   Excelデータの1行分の情報を保持し、セルデータの追加、取得、およびイテレータを提供する
#
############################################################
class LineData():
    def __init__(self, lineData):
        self.data = lineData

    def getCell(self, col):
        return self.data[col - 1]

    def setCell(self, col, data):
        self.data[col - 1] = data

    def append(self, data):
        self.data.append(data)

    def __iter__(self):
        self.itrCount = 0
        return self

    def __next__(self):
        if self.itrCount >= len(self.data):
            raise StopIteration()
        self.itrCount += 1
        return self.data[self.itrCount - 1]

