# -*- coding: utf-8 -*-
from Debug import Debug
from LineData import LineData
from ExcelMan import ExcelMan


############################################################
#
#   1ブックのExcelデータを保持する
#   データの書き込みを提供する
#
############################################################

class BookData(ExcelMan):
    """
       1ブックのExcelデータを保持する
       データの書き込みを提供する
    """
    def __init__(self, file):
        super().__init__(file)
        super().openBook()

    def openSheetByName(self, sheetName):
        """
        シートをオープンする
        :param sheetName: シート名
        :return:
        """
        try:
            super().openSheetByName(sheetName)  # 対象となるシートを開く
        except:
            print(sheetName+"シートが開けません")
            raise
