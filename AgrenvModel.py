# -*- coding: utf-8 -*-
from openpyxl.cell import Cell

from Debug import Debug

from TitleInfo import TitleInfo
from SheetData import SheetData
from Approach import Approach
import re
import copy

class AgrenvModel():
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
        self.PRODUCE_TYPE = TitleInfo(r"作物区分")
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
            self.PRODUCE_TYPE,
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

        self.periodList = []            # 実施期間、栽培期間のリスト
        self.areaList = []              # 構成員別取組面積のリスト

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
    #   Excelのデータを読み込み、今回の処理に必要な情報をピックアップして
    #   内部形式に変換して、self.approachsのリストに保持する
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
        self.PRODUCE_TYPE.searchOnLine(lineData)  # 作物区分
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
            self.PRODUCE_TYPE,  # 作物区分
            self.CULTIVATED_START_YEAR,  # 栽培開始年
            self.CULTIVATED_START_MONTH,  # 栽培開始月
            self.CULTIVATED_END_YEAR,  # 栽培終了年
            self.CULTIVATED_END_MONTH  # 栽培終了月
        )

        self.pickupApproachList()
        self.approachProc()
        self.areaListProc()
        self.sum()
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

    #############################
    #
    #   取り組み一覧関係
    #   取り組み名称毎に、作物に対する、実施期間と栽培期間のリストを作る
    #
    #############################
    def approachProc(self):
        """
        取り組み一覧関係
        取り組み名称毎に、作物に対する、実施期間と栽培期間のリストを作る
        :return:
        """
        duplicateList = copy.copy(self.approachs)         # リストの内容を変更するので、複製を作成する
        duplicateList.sort(key=Approach.getApproachAndProduce)
        Debug.print("=============ソート完了=============")
        for app in duplicateList:
            Debug.print( app.print() )

        compareName = None              # 比較する名前
        self.periodList = []               # 重複を取り除いたリスト
        for app in duplicateList:
            if app.getApproachAndProduce() == compareName:
                continue
            self.periodList.append(app)
            compareName = app.getApproachAndProduce()

        Debug.print("=============重複削除完了=============")
        for app in self.periodList:
            Debug.print( app.print() )

    #############################
    #
    #   構成員別取り組み面積
    #
    #############################
    def areaListProc(self):
        """
        構成員別取り組み面積
        :return:
        """
        duplicateList = copy.copy(self.approachs)         # リストの内容を変更するので、複製を作成する
        duplicateList.sort(key=Approach.getPersonAndApproach)
        Debug.print("=============ソート完了=============")
        for app in duplicateList:
            Debug.print( app.print() )

        compareName = None              # 比較する名前
        self.areaList = []               # 重複を取り除いたリスト
        for app in duplicateList:
            if app.getPersonAndApproach() == compareName:
                areaSum["area"] += app.area
                continue
            areaSum = {}
            areaSum["approach"] = app.approachType
            areaSum["area"] = app.area

            self.areaList.append(areaSum)
            compareName = app.getPersonAndApproach()

        Debug.print("=============重複削除完了=============")
        for app in self.areaList:
           print( app  )

    #############################
    #
    #   集計
    #
    #############################
    def sum(self):
        """
        取組毎の面積を集計する
        :return:
        """
        duplicateList = copy.copy(self.approachs)         # リストの内容を変更するので、複製を作成する
        duplicateList.sort(key=Approach.getApproachType)
        Debug.print("=============ソート完了=============")
        for app in duplicateList:
            Debug.print( app.print() )

        compareName = None              # 比較する名前
        self.sumList = []               # 重複を取り除いたリスト
        for app in duplicateList:
            if app.getApproachType() == compareName:
                areaSum["area"] += app.area
                continue
            areaSum = {}
            areaSum["approach"] = app.approachType
            areaSum["area"] = app.area

            self.sumList.append(areaSum)
            compareName = app.getApproachType()

        Debug.print("=============重複削除完了=============")
        for app in self.sumList:
           print( app  )

###########################
#   テスト
###########################
if __name__ == '__main__':
    print("####テストスタート#####")
    tergetObj = AgrenvModel("実施計画書(元データ)2.xlsx")  # "sample.xlsxファイルの管理オブジェクトを作る
