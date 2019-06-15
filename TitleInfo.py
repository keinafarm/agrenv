# -*- coding: utf-8 -*-
import re
from Debug import Debug

############################################################
#
#   表題の情報を管理するクラス
#   表題に含まれる文字（正規表現）を指定しsearchOnLineを実行すると
#   該当する文字を探してカラム番号を覚えておく
#
############################################################

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

