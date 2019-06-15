# -*- coding: utf-8 -*-
from openpyxl.cell import Cell

from ExcelMan import ExcelMan
from Debug import Debug
from EraProc import EraProc
import re


class TitleInfo():
    """
    タイトル行の情報
    """

    def __init__(self, pattern):
        """
        コンストラクタ
        :param pattern: 検索パターン
        """
        self.pattern = pattern
        self.cellObj = None  # シート上で、パターンに一致したセルのオブジェクト
        self.colNum = 0

    def searchOnLine(self, lineData):
        """
        指定した行のセルの中に特定のパターンを含む最初のセルを見つける
        :param lineNo: 検索する行番号(1オリジン)
        :return:　セルオブジェクト　無い時はNone
        """
        for colData in lineData:  # 行の先頭からチェック
            data = colData.value  # セルのデータを取得

            if data == None:  # データが無い
                continue

            if not (type(data) is str):  # 文字列でない
                continue

            result = re.search(self.pattern, colData.value)  # 取組名称の文字を探す

            if result:  # 見つかったらループ終了
                self.cellObj = colData
                Debug.print(
                    self.pattern + " row=" + str(self.cellObj.row) + " col=" + str(self.cellObj.column))
                self.colNum = self.cellObj.column
                return True

        self.colNum = None
        Debug.error(self.pattern + "のタイトルが見つかりませんでした")  # Debug用
        return False

    def col(self):
        return self.colNum


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


class ExcelManForAgrenv():
    """
    環境直払い　実施計画書のファイルを読み込んで、申請書類のデータを作成する
    """

    TERGET_SHEET_NAME = "◎ほ場一覧（全構成員）"  # 対象となるシートの名前

    #############################
    #
    #   実施計画書ファイルに対する集計処理
    #
    #############################

    def __init__(self, file):
        """
        実施計画書ファイルを開く
        :param file: 実施計画書のファイル名
        """

        self.data = None  # ほ場一覧シートのデータを保持する
        self.dataLine = 0  # データ開始行
        # タイトル情報
        self.PERSON_NAME = TitleInfo(r"構成員(.|\n)*?（漢字）")
        self.AREA = TitleInfo(r"取組面積")
        self.TITLE_LINE_KEY = TitleInfo(r"取組名称")
        self.IMPLE_START_YEAR = TitleInfo(r"実施(.|\n)*?時期(.|\n)*?開始年")
        self.IMPLE_START_MONTH = TitleInfo(r"実施(.|\n)*?時期(.|\n)*?開始月")
        self.IMPLE_END_YEAR = TitleInfo(r"実施(.|\n)*?時期(.|\n)*?終了年")
        self.IMPLE_END_MONTH = TitleInfo(r"実施(.|\n)*?時期(.|\n)*?終了月")
        self.PRODUCE_NAME = TitleInfo(r"作物名")
        self.CULTIVATED_START_YEAR = TitleInfo(r"栽培(.|\n)*?時期(.|\n)*?開始年")
        self.CULTIVATED_START_MONTH = TitleInfo(r"栽培(.|\n)*?時期(.|\n)*?開始月")
        self.CULTIVATED_END_YEAR = TitleInfo(r"栽培(.|\n)*?時期(.|\n)*?開始年")
        self.CULTIVATED_END_MONTH = TitleInfo(r"栽培(.|\n)*?時期(.|\n)*?終了月")

        # タイトル情報を保持する
        titleCellList = [
            self.TITLE_LINE_KEY,
            self.PERSON_NAME,
            self.AREA,
            self.IMPLE_START_YEAR,
            self.IMPLE_START_MONTH,
            self.IMPLE_END_YEAR,
            self.IMPLE_END_MONTH,
            self.PRODUCE_NAME,
            self.CULTIVATED_START_YEAR,
            self.CULTIVATED_START_MONTH,
            self.CULTIVATED_END_YEAR,
            self.CULTIVATED_END_MONTH
        ]

        self.data = SheetData(file, self.TERGET_SHEET_NAME)  # シートデータを読み込む
        self.approachs = []

        if not (self.approachList()):
            Debug.error("データの形式が思ってた通りでは無かったです　ごめんなさい")
            raise ValueError

    #############################
    #
    #   クラス内汎用処理
    #
    #############################

    def searchCells(self, pattern):
        """
        特定のパターンを含む最初のセルを見つける
        :param pattern:　正規表現で記載された検索パターン
        :return:　セルオブジェクト　無い時はNone
        """
        for lineNo in range(1, self.data.numOfRow()):  # 最初の行からチェック
            result = self.searchInLine(lineNo, pattern)
            if result:
                return result

        return None

    def searchInLine(self, rowNo, pattern):
        """
        指定した行番号のセルの中に特定のパターンを含む最初のセルを見つける
        :param rowNo: 検索する行番号(1オリジン)
        :return:　セルオブジェクト　無い時はNone
        """
        colData: Cell
        lineData = self.data.selectline(rowNo)
        for cellData in lineData:  # 行の先頭からチェック
            data = cellData.value  # セルのデータを取得

            if data == None:  # データが無い
                continue

            if not (type(data) is str):  # 文字列でない
                continue

            Debug.print(cellData.value)  # Debug用
            result = re.search(pattern, cellData.value)  # 取組名称の文字を探す

            if result:  # 見つかったらループ終了
                return cellData

        return None

    #############################
    #
    #   取り組み一覧関係
    #   各取り組みと作物の実施時期と栽培時期を抽出する
    #
    #############################
    def approachList(self):
        """
        取り組み一覧を作成する
        取り組み一覧は
            (8) 取組名称 (１取組目)
            (9) 実施 時期 開始年
            (10) 実施 時期 開始月
            (11) 実施 時期 終了年
            (12) 実施 時期 終了月
            (13) 作物区分 (１取組目)
            (14) 作物名 (１取組目)
            (15) 栽培 時期 開始年
            (16) 栽培 時期 開始月
            (17) 栽培 時期 終了年
            (18) 栽培 時期 終了月
        の列からデータを取得する
        :return True 作成成功
                False　作成失敗
        """

        print("タイトル行の検索開始")
        obj = self.searchCells(self.TITLE_LINE_KEY.pattern)
        if obj is None:
            Debug.error('"取組名称"を含むセルが見つかりません')
            return False

        Debug.print("Find row=" + str(obj.row) + " col=" + str(obj.column))
        titleLine = obj.row  # 表題の行
        self.dataLine = titleLine + 1  # データの開始行
        lineData = self.data.selectline(titleLine)

        self.TITLE_LINE_KEY.searchOnLine(lineData)  # 取組名称
        self.PERSON_NAME.searchOnLine(lineData)  # 構成員名
        self.AREA.searchOnLine(lineData)  # 取り組み面積
        self.IMPLE_START_YEAR.searchOnLine(lineData)  # 実施開始年
        self.IMPLE_START_MONTH.searchOnLine(lineData)  # 実施開始月
        self.IMPLE_END_YEAR.searchOnLine(lineData)  # 実施終了年
        self.IMPLE_END_MONTH.searchOnLine(lineData)  # 実施終了月
        self.PRODUCE_NAME.searchOnLine(lineData)  # 作物名
        self.CULTIVATED_START_YEAR.searchOnLine(lineData)  # 栽培開始年
        self.CULTIVATED_START_MONTH.searchOnLine(lineData)  # 栽培開始月
        self.CULTIVATED_END_YEAR.searchOnLine(lineData)  # 栽培終了年
        self.CULTIVATED_END_MONTH.searchOnLine(lineData)  # 栽培終了月
        Approach.setInfoLocation(
            self.TITLE_LINE_KEY,  # 取組名称
            self.PERSON_NAME,  # 構成員名
            self.AREA,  # 構成員名
            self.IMPLE_START_YEAR,  # 実施開始年
            self.IMPLE_START_MONTH,  # 実施開始月
            self.IMPLE_END_YEAR,  # 実施終了年
            self.IMPLE_END_MONTH,  # 実施終了月
            self.PRODUCE_NAME,  # 作物名
            self.CULTIVATED_START_YEAR,  # 栽培開始年
            self.CULTIVATED_START_MONTH,  # 栽培開始月
            self.CULTIVATED_END_YEAR,  # 栽培終了年
            self.CULTIVATED_END_MONTH  # 栽培終了月
        )

        self.pickupApproachList()
        return True

    def pickupApproachList(self):
        """
        取り組み一覧の対象となるデータを抽出する
        :return:
        """
        self.approachs = []
        for lineNo in range(self.dataLine, self.data.numOfRow()):
            approachLine = Approach.factoryApproach(self.data.selectline(lineNo))
            if approachLine is None:
                continue
            self.approachs.append(approachLine)
            Debug.print(approachLine.print())


class Approach():
    """
    取り組み対象を処理するクラス
    """
    APPROACH_NAMES = [
        "●対象外",
        "カバークロップ",
        "堆肥施用",
        "有機農業",
        "冬季湛水",
    ]
    TITLE_LINE_KEY = None  # 取組名称
    PERSON_NAME = None  # 構成員名
    AREA = None  # 取り組み面積
    IMPLE_START_YEAR = None  # 実施開始年
    IMPLE_START_MONTH = None  # 実施開始月
    IMPLE_END_YEAR = None  # 実施終了年
    IMPLE_END_MONTH = None  # 実施終了月
    PRODUCE_NAME = None  # 作物名
    CULTIVATED_START_YEAR = None  # 栽培開始年
    CULTIVATED_START_MONTH = None  # 栽培開始月
    CULTIVATED_END_YEAR = None  # 栽培終了年
    CULTIVATED_END_MONTH = None  # 栽培終了月

    @classmethod
    def setInfoLocation(cls,
                        TITLE_LINE_KEY,  # 取組名称
                        PERSON_NAME,  # 構成員名
                        AREA,  # 取り組み面積
                        IMPLE_START_YEAR,  # 実施開始年
                        IMPLE_START_MONTH,  # 実施開始月
                        IMPLE_END_YEAR,  # 実施終了年
                        IMPLE_END_MONTH,  # 実施終了月
                        PRODUCE_NAME,  # 作物名
                        CULTIVATED_START_YEAR,  # 栽培開始年
                        CULTIVATED_START_MONTH,  # 栽培開始月
                        CULTIVATED_END_YEAR,  # 栽培終了年
                        CULTIVATED_END_MONTH  # 栽培終了月
                        ):
        """
        各情報の位置を決定するための情報を受取る
        :param TITLE_LINE_KEY:
        :param PERSON_NAME:
        :param AREA:
        :param IMPLE_START_YEAR:
        :param IMPLE_START_MONTH:
        :param IMPLE_END_YEAR:
        :param IMPLE_END_MONTH:
        :param PRODUCE_NAME:
        :param CULTIVATED_START_YEAR:
        :param CULTIVATED_START_MONTH:
        :param CULTIVATED_END_YEAR:
        :param CULTIVATED_END_MONTH:
        :return:
        """
        cls.TITLE_LINE_KEY = TITLE_LINE_KEY  # 取組名称
        cls.PERSON_NAME = PERSON_NAME  # 構成員名
        cls.AREA = AREA  # 取り組み面積
        cls.IMPLE_START_YEAR = IMPLE_START_YEAR  # 実施開始年
        cls.IMPLE_START_MONTH = IMPLE_START_MONTH  # 実施開始月
        cls.IMPLE_END_YEAR = IMPLE_END_YEAR  # 実施終了年
        cls.IMPLE_END_MONTH = IMPLE_END_MONTH  # 実施終了月
        cls.PRODUCE_NAME = PRODUCE_NAME  # 作物名
        cls.CULTIVATED_START_YEAR = CULTIVATED_START_YEAR  # 栽培開始年
        cls.CULTIVATED_START_MONTH = CULTIVATED_START_MONTH  # 栽培開始月
        cls.CULTIVATED_END_YEAR = CULTIVATED_END_YEAR  # 栽培終了年
        cls.CULTIVATED_END_MONTH = CULTIVATED_END_MONTH  # 栽培終了月

    @classmethod
    def factoryApproach(cls, lineData):
        # 取組名称の種別をチェックする
        cell = lineData.getCell(cls.TITLE_LINE_KEY.col()).value
        if not (type(cell) is str):
            return None  # 文字列ではないので無視

        for typeName in cls.APPROACH_NAMES:
            result = re.search(cell, typeName)
            if result:
                break;

        retObj = Approach(lineData, typeName)
        if (result == None) or (typeName == cls.APPROACH_NAMES[0]):
            return None  # 一致しない取組名称と、●対象外は無視

        try:
            # 構成員名を得る
            cell = lineData.getCell(Approach.PERSON_NAME.col()).value
            retObj.setPersonName(cell)

            # 面積を得る
            cell = lineData.getCell(Approach.AREA.col()).value
            area = float(cell)
            retObj.setArea(area)

            # 実施開始期間を得る
            cell = lineData.getCell(Approach.IMPLE_START_YEAR.col()).value
            year = int(cell)
            cell = lineData.getCell(Approach.IMPLE_START_MONTH.col()).value
            month = int(cell)
            start = EraProc(year, month, 1)

            # 実施終了期間を得る
            cell = lineData.getCell(Approach.IMPLE_END_YEAR.col()).value
            year = int(cell)
            cell = lineData.getCell(Approach.IMPLE_END_MONTH.col()).value
            month = int(cell)
            end = EraProc(year, month, 1)
            retObj.setImple(start, end)

            # 栽培開始期間を得る
            cell = lineData.getCell(Approach.CULTIVATED_START_YEAR.col()).value
            year = int(cell)
            cell = lineData.getCell(Approach.CULTIVATED_START_MONTH.col()).value
            month = int(cell)
            start = EraProc(year, month, 1)

            # 栽培終了期間を得る
            cell = lineData.getCell(Approach.CULTIVATED_END_YEAR.col()).value
            year = int(cell)
            cell = lineData.getCell(Approach.CULTIVATED_END_MONTH.col()).value
            month = int(cell)
            end = EraProc(year, month, 1)
            retObj.setculti(start, end)

            name = lineData.getCell(Approach.PRODUCE_NAME.col()).value
            retObj.setProduce(name)
        except:
            Debug.error("期間が不正です。無視します")
            return None

        Debug.print(retObj.print())
        return retObj

    def __init__(self, lineData, typeName):
        """
        コンストラクタ
        :param lineData: １行分のデータ
        """
        self.apptoachType = typeName
        self.data = lineData

    def setPersonName(self, name):
        self.personName = name

    def setArea(self, area):
        self.area = area

    def setImple(self, start, end):
        self.impleStart = start
        self.impleEnd = end

    def setculti(self, start, end):
        self.cultiStart = start
        self.cultiEnd = end

    def setProduce(self, name):
        self.produceName = name

    def print(self):
        return self.personName + " (" + str(self.area) + "a)" + \
               self.apptoachType + ":" + self.produceName + " = " + self.impleStart.print() + "～" + self.impleEnd.print() + \
               " |  " + self.cultiStart.print() + "～" + self.cultiEnd.print()


###########################
#   テスト
###########################
if __name__ == '__main__':
    print("####テストスタート#####")
    tergetObj = ExcelManForAgrenv("実施計画書(元データ)2.xlsx")  # "sample.xlsxファイルの管理オブジェクトを作る
