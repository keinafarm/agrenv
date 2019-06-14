# -*- coding: utf-8 -*-
from openpyxl.cell import Cell

from ExcelMan import ExcelMan
from Debug import Debug
import re

class   titleInfo():
    """
    タイトル行の情報
    """
    pattern = None  # どのタイトルなのか検索する時のパターン
    cellObj = None   # シート上で、パターンに一致したセルのオブジェクト

    def __init__(self, pattern):
        """
        コンストラクタ
        :param pattern: 検索パターン
        """
        self.pattern = pattern

    def searchOnLine(self, lineData ):
        """
        指定した行のセルの中に特定のパターンを含む最初のセルを見つける
        :param lineNo: 検索する行番号(1オリジン)
        :return:　セルオブジェクト　無い時はNone
        """
        for colData in lineData:            # 行の先頭からチェック
            data = colData.value  # セルのデータを取得

            if data == None:            #データが無い
                continue

            if not(type(data) is str):      # 文字列でない
                continue

            result = re.search(self.pattern, colData.value)  # 取組名称の文字を探す

            if result:                                  # 見つかったらループ終了
                self.cellObj = colData
                Debug.print(self.pattern + " row=" + str(self.cellObj.row) + " col="+ str( self.cellObj.column) )
                return True

        self.colNum = None
        Debug.error( self.pattern +  "のタイトルが見つかりませんでした")     # Debug用
        return False


class ExcelManForAgrenv(ExcelMan):
    """
    環境直払い　実施計画書のファイルを読み込んで、申請書類のデータを作成する
    """

    rows = 0            # 行数
    cols = 0            # 桁数
    data = []           # ほ場一覧シートのデータを保持する

    # タイトル情報
    TITLE_LINE_KEY = titleInfo(r"取組名称")
    IMPLE_START_YEAR = titleInfo(r"実施(.|\n)*?時期(.|\n)*?開始年")
    IMPLE_START_MONTH = titleInfo(r"実施(.|\n)*?時期(.|\n)*?開始月")
    IMPLE_END_YEAR = titleInfo(r"実施(.|\n)*?時期(.|\n)*?終了年")
    IMPLE_END_MONTH = titleInfo(r"実施(.|\n)*?時期(.|\n)*?終了月")
    PRODUCE_NAME = titleInfo(r"作物名")
    CULTIVATED_START_YEAR = titleInfo(r"栽培(.|\n)*?時期(.|\n)*?開始年")
    CULTIVATED_START_MONTH = titleInfo(r"栽培(.|\n)*?時期(.|\n)*?開始月")
    CULTIVATED_END_YEAR = titleInfo(r"栽培(.|\n)*?時期(.|\n)*?開始年")
    CULTIVATED_END_MONTH = titleInfo(r"栽培(.|\n)*?時期(.|\n)*?終了月")

    # タイトル情報を保持する
    titleCellList = [
        TITLE_LINE_KEY,
        IMPLE_START_YEAR,
        IMPLE_START_MONTH,
        IMPLE_END_YEAR,
        IMPLE_END_MONTH,
        PRODUCE_NAME,
        CULTIVATED_START_YEAR,
        CULTIVATED_START_MONTH,
        CULTIVATED_END_YEAR,
        CULTIVATED_END_MONTH
    ]

    def __init__(self, file):
        """
        実施計画書ファイルを開く
        :param file: 実施計画書のファイル名
        """
        super().__init__(file)
        super().openBook()
        self.readAll()
        if not (self.approachList()):
            Debug.error( "データの形式が思ってた通りでは無かったです　ごめんなさい")
            raise ValueError

    def readAll(self):
        """
        実施計画書ファイルの、"◎ほ場一覧（全構成員）"シートを全部読み込む
        :return:
        """
        super().openSheetByName("◎ほ場一覧（全構成員）")
        self.rows = super().numOfRow()
        self.cols = super().numOfCol()

        print( "◎ほ場一覧（全構成員）を読み込み 行数="+str(self.rows) + " 列数="+str(self.cols))
        self.data = []
        for row in range(1,self.rows):
            line = []
            for col in range(1,self.cols):
                cellObj = super().getCell( row,col)
                line.append(cellObj)
            self.data.append(line)

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

        Debug.print( "Find row="+str(obj.row) + " col="+ str(obj.column))
        titleLine = obj.row

        self.TITLE_LINE_KEY.searchOnLine(self.data[titleLine-1])        # 取組名称
        self.IMPLE_START_YEAR.searchOnLine(self.data[titleLine-1])       # 実施開始年
        self.IMPLE_START_MONTH.searchOnLine(self.data[titleLine-1])       # 実施開始月
        self.IMPLE_END_YEAR.searchOnLine(self.data[titleLine-1])       # 実施終了年
        self.IMPLE_END_MONTH.searchOnLine(self.data[titleLine-1])       # 実施終了月
        self.PRODUCE_NAME.searchOnLine(self.data[titleLine-1])           # 作物名
        self.CULTIVATED_START_YEAR.searchOnLine(self.data[titleLine-1])  # 栽培開始年
        self.CULTIVATED_START_MONTH.searchOnLine(self.data[titleLine-1])  # 栽培開始月
        self.CULTIVATED_END_YEAR.searchOnLine(self.data[titleLine-1])  # 栽培終了年
        self.CULTIVATED_END_MONTH.searchOnLine(self.data[titleLine-1])  # 栽培終了月
        return True

    def searchCells(self, pattern):
        """
        特定のパターンを含む最初のセルを見つける
        :param pattern:　正規表現で記載された検索パターン
        :return:　セルオブジェクト　無い時はNone
        """
        for lineNo in range(1,self.numOfRow()):           # 最初の行からチェック
            result = self.searchInLine(lineNo, pattern)
            if result:
                return result

        return None

    def searchInLine(self, lineNo, pattern):
        """
        指定した行番号のセルの中に特定のパターンを含む最初のセルを見つける
        :param lineNo: 検索する行番号(1オリジン)
        :return:　セルオブジェクト　無い時はNone
        """
        colData: Cell
        lineData = self.data[lineNo-1]
        for colData in lineData:            # 行の先頭からチェック
            data = colData.value  # セルのデータを取得

            if data == None:            #データが無い
                continue

            if not(type(data) is str):      # 文字列でない
                continue

            Debug.print(colData.value)            # Debug用
            result = re.search(pattern, colData.value)  # 取組名称の文字を探す

            if result:                                  # 見つかったらループ終了
                return colData

        return None

###########################
#   テスト
###########################
if __name__ == '__main__':
    print("####テストスタート#####")
    tergetObj = ExcelManForAgrenv("実施計画書(元データ)2.xlsx")             # "sample.xlsxファイルの管理オブジェクトを作る
    tergetObj.readAll()
    tergetObj.approachList()
