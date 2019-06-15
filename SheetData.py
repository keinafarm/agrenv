# -*- coding: utf-8 -*-
from Debug import Debug
from LineData import LineData
from ExcelMan import ExcelMan


############################################################
#
#   1シートのExcelデータを保持する
#   Excelデータの1シートの情報を保持し、
#   処理中の行を覚えて、提供する
#
############################################################

class SheetData(ExcelMan):
    """
    シートのデータを管理する
    """

    def __init__(self, file, sheet):
        self._rows = 0  # 行数
        self._cols = 0  # 桁数
        self._data = []  # ほ場一覧シートのデータを保持する
        self._lineData = []  # １行分のデータ
        self._currentRow = 0  # 現在着目している行番号（カレント行）

        super().__init__(file)
        super().openBook()
        super().openSheetByName(sheet)  # 対象となるシートを開く
        self._rows = super().numOfRow()
        self._cols = super().numOfCol()
        print(sheet + "を読み込み 行数=" + str(self._rows) + " 列数=" + str(self._cols))
        lineData = LineData([])
        for row in range(1, self._rows):
            for col in range(1, self._cols):
                cellObj = super().getCell(row, col)
                lineData.append(cellObj)
            self._data.append(lineData)
            lineData = LineData([])

    def selectline(self, rowNo):
        """
        1行分のデータを得る
        :param rowNo: 行番号
        :return:
        """
        self._lineData = self._data[rowNo - 1]
        self._currentRow = rowNo
        return self._lineData

    def getLine(self):
        """
        カレント行を得る
        :return:
        """
        return self.selectline(self._currentRow)

    def getCell(self, row, col):
        return self._data[row - 1].getCell(col).value

    def getCellOnCurrentLine(self, col):
        return self._lineData.getCell(col).value

    def numOfRow(self):
        return self._rows

    def numOfCol(self):
        return self._cols