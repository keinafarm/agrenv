# -*- coding: utf-8 -*-
from Debug import Debug
from EraProc import EraProc
import re

############################################################
#
#   1行分の取り組み情報を管理する
#   Excelのデータを内部表現に変換して保持する
#
############################################################
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
    APPROACH_NAME = None  # 取組名称
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
                        APPROACH_NAME,  # 取組名称
                        PERSON_NAME,  # 構成員名
                        AREA,  # 取り組み面積
                        IMPLE_START_YEAR,  # 実施開始年
                        IMPLE_START_MONTH,  # 実施開始月
                        IMPLE_END_YEAR,  # 実施終了年
                        IMPLE_END_MONTH,  # 実施終了月
                        PRODUCE_NAME,  # 作物名
                        PRODUCE_TYPE,   # 作物区分
                        CULTIVATED_START_YEAR,  # 栽培開始年
                        CULTIVATED_START_MONTH,  # 栽培開始月
                        CULTIVATED_END_YEAR,  # 栽培終了年
                        CULTIVATED_END_MONTH  # 栽培終了月
                        ):
        """
        各情報の位置を決定するための情報を受取る
        :param APPROACH_NAME:
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
        cls.APPROACH_NAME = APPROACH_NAME  # 取組名称
        cls.PERSON_NAME = PERSON_NAME  # 構成員名
        cls.AREA = AREA  # 取り組み面積
        cls.IMPLE_START_YEAR = IMPLE_START_YEAR  # 実施開始年
        cls.IMPLE_START_MONTH = IMPLE_START_MONTH  # 実施開始月
        cls.IMPLE_END_YEAR = IMPLE_END_YEAR  # 実施終了年
        cls.IMPLE_END_MONTH = IMPLE_END_MONTH  # 実施終了月
        cls.PRODUCE_NAME = PRODUCE_NAME  # 作物名
        cls.PRODUCE_TYPE = PRODUCE_TYPE  # 作物区分
        cls.CULTIVATED_START_YEAR = CULTIVATED_START_YEAR  # 栽培開始年
        cls.CULTIVATED_START_MONTH = CULTIVATED_START_MONTH  # 栽培開始月
        cls.CULTIVATED_END_YEAR = CULTIVATED_END_YEAR  # 栽培終了年
        cls.CULTIVATED_END_MONTH = CULTIVATED_END_MONTH  # 栽培終了月

    @classmethod
    def factoryApproach(cls, lineData):
        # 取組名称の種別をチェックする
        cell = lineData.getCell(cls.APPROACH_NAME.col()).value
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

            #　作物名を得る
            name = lineData.getCell(Approach.PRODUCE_NAME.col()).value
            retObj.setProduce(name)

            #　作物区分を得る
            name = lineData.getCell(Approach.PRODUCE_TYPE.col()).value
            retObj.setProduceType(name)

        except:
            Debug.error("書式が不正です。無視します" )
            return None
        return retObj

    def __init__(self, lineData, typeName):
        """
        コンストラクタ
        :param lineData: １行分のデータ
        """
        self.approachType = typeName
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

    def setProduceType(self, name):
        self.produceType = name

    # ソート用
    def getProduce(self):
        return  self.produceName

    def getApproachAndProduce(self):
        """
        取り組み毎に、作物名,でソートする
        :return:
        """
        return self.approachType + "$$$$$$$" + self.produceName

    def getPersonAndApproach(self):
        """
        構成員毎に、取組名称、作物区分でソートする
        :return:
        """
        return self.personName + "$$$$$$$" + self.approachType + "$$$$$$$" + self.produceType

    def getApproachType(self):
        return  self.approachType

    def print(self):
        return self.personName + " (" + str(self.area) + "a)" + \
               self.approachType + ":" + self.produceName + "(" + self.produceType + ") = " + self.impleStart.print() + "～" + self.impleEnd.print() + \
               " |  " + self.cultiStart.print() + "～" + self.cultiEnd.print()

