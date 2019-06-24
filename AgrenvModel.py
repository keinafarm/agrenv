# -*- coding: utf-8 -*-
from openpyxl.cell import Cell

from Debug import Debug

from TitleInfo import TitleInfo
from SheetData import SheetData
from BookData import BookData
from Approach import Approach
import re
import copy
import math

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
        self.periodList = []            # 実施期間、栽培期間のリスト
        self.areaList = []              # 構成員別取組面積のリスト

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
        print("********** 取り組み一覧*************")
        for app in self.periodList:
            print( app.print() )

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
        areaSum = {}
        for app in duplicateList:
            if app.getPersonAndApproach() == compareName:
                areaSum["area"] += app.area
                continue
            areaSum = {}
            areaSum["personName"] = app.personName
            areaSum["approach"] = app.approachType
            areaSum["produceType"] = app.produceType
            areaSum["area"] = app.area

            self.areaList.append(areaSum)
            compareName = app.getPersonAndApproach()

        Debug.print("=============重複削除完了=============")
        print("**********構成員別取り組み面積*************")
        for app in self.areaList:
           print( app )

    #############################
    #
    #   交付申請書用　取組毎の面積を集計
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
        areaSum = {}
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
        print("**********取組毎の面積*************")
        for app in self.sumList:
           print( app )

    #############################
    #
    #   出力
    #
    #############################
    SUM_SHEET_NAME = "◎集計(実施計画書から）"
    INITIATIVES_SHEET_NAME = "◎取り組み一覧（実施計画書から)"

    def output(self, applicationFile, planFile):
        """
        交付申請書と実施計画書ファイルに出力する
        :param applicationFile: 交付申請書ファイル名
        :param planFile: 事業計画書ファイル名
        :return:
        """
        try:
            self.appliBook = BookData(applicationFile)
            self.planBook = BookData(planFile)

#            self.appliInitiativesSheet = self.appliBook.tergetSheet(self.INITIATIVES_SHEET_NAME)  # 申請書：取組
#            self.planInitiativesSheet = self.planBook.tergetSheet(self.INITIATIVES_SHEET_NAME)  # 事業計画書：集計シート

        except OSError as err:
            print("OS error: {0}".format(err))

        self.sumOut( self.appliBook)
        self.sumOut( self.planBook)
        self.appInitiativesOut( self.appliBook)
        self.planpInitiativesOut( self.planBook)

    def sumOut(self, book):
        """
        集計シートに出力する
        :param book: 出力するブックオブジェクト
        :return:
        """
        ### 定形パターンを記載する
        fixedPattern = [
            ["", "", "", "補助金単価", "1取り組み面積", "", "1取り組み金額", ""],
            ["", "カバークロップ", "ヒエの種子", 7000, 0, "=sum(E2)", "=D2*F2/10", "=D2*F2/10"],
            ["", "カバークロップ", "ヒエ以外", 8000, 0, "=sum(E3)", "=D3*F3/10", "=D3*F3/10"],
            ["", "堆肥施用", "", 4400, 0, "=sum(E4)", "=D4*F4/10", "=D4*F4/10"],
            ["", "有機農業", "", 8000, 0, "=sum(E5)", "=D5*F5/10", "=D5*F5/10"],
            ["", "冬期湛水", "1", 8000, 0, "=sum(E6)", "=D6*F6/10", "=D6*F6/10"],
            ["", "冬期湛水", "2", 7000, 0, "=sum(E7)", "=D7*F7/10", "=D7*F7/10"],
            ["", "冬期湛水", "3", 5000, 0, "=sum(E8)", "=D8*F8/10", "=D8*F8/10"],
            ["", "冬期湛水", "4", 4000, 0, "=sum(E9)", "=D9*F9/10", "=D9*F9/10"],
            ["", "インセクタリープランツ", "", 8000, 0, "=sum(E10)", "=D10*F10/10", "=D10*F10/10"],
            ["", "", "", "", "=sum(E2:E10)", "=sum(F2:F10)", "=sum(G2:G10)", "=sum(H2:H10)"],
        ]
        book.openSheetByName(self.SUM_SHEET_NAME)  # 集計シート
        for row in range(1,12):
            for col in range(1,9):
                book.setCell(row, col, fixedPattern[row-1][col-1])

        for item in self.sumList:
            if item["approach"] == "カバークロップ":
                book.setCell(3, 5,math.floor(item["area"]))
            elif item["approach"] == "堆肥施用":
                book.setCell(4, 5,math.floor(item["area"]))
            elif item["approach"] == "有機農業":
                book.setCell(5, 5,math.floor(item["area"]))
            elif item["approach"] == "冬期湛水":
                book.setCell(7, 5,math.floor(item["area"]))

        book.saveBook()

    def appInitiativesOut(self, book):
        """
        申請書：取り組み一覧出力
        :param book:出力するブックオブジェクト
        :return:
        """
        book.openSheetByName(self.INITIATIVES_SHEET_NAME)  # 取組一覧シート
        # クリア
        for row in range(1,12):
            for col in range(1,5):
                book.setCell(row, col, "")

        #　取り組み一覧出力をセルに格納する
        for row in range(1,len(self.areaList)+1):
            book.setCell(row, 1, self.areaList[row-1]["personName"])
            book.setCell(row, 2, self.areaList[row-1]["approach"])
            book.setCell(row, 3, self.areaList[row-1]["produceType"])
            book.setCell(row, 4, self.areaList[row-1]["area"])

        book.saveBook()

    def planpInitiativesOut(self, book):
        """
        事業計画書:取り組み一覧
        :param book: 事業計画書ファイルオブジェクト
        :return:
        """
        book.openSheetByName(self.INITIATIVES_SHEET_NAME)  # 取組一覧シート
        # クリア
        for row in range(1,20):
            for col in range(1,5):
                book.setCell(row, col, "")

        row = 1
        for item in self.periodList:
            book.setCell(row, 1, item.produceName)
            impleString = (item.impleStart.month()) + "月～" +   (item.impleEnd.month()) + "月"
            book.setCell(row, 2, impleString)
            book.setCell(row, 3, item.produceName)
            cultiString = (item.cultiStart.month()) + "月～" +   (item.cultiEnd.month()) + "月"
            book.setCell(row, 4, cultiString)
            row += 1

        book.saveBook()

###########################
#   テスト
###########################
if __name__ == '__main__':
    print("####テストスタート#####")
    tergetObj = AgrenvModel("実施計画書(元データ)2.xlsx")  # "sample.xlsxファイルの管理オブジェクトを作る

    tergetObj.output("testA.xlsx", "testP.xlsx")

