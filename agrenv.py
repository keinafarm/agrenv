# -*- coding: utf-8 -*-
from ArgMan import ArgMan
from ExcelMan import ExcelMan
import re

#################################################################
#
#   環境保全型農業直接支払い交付金の為のデータ作成
#
#################################################################
#
class   agrenv:
    outFile = "out.xlsx"
    dir = "."
    file = '実施計画書(元データ).xlsx'

    def __init__(self):
        self.__argMan = ArgMan("環境保全型農業直接支払い交付金の実施計画書からデータを抜き出す",
                        [
                            ["--o", None, self.outFile, None, '出力ファイル名', '出力ファイル指定(指定しない場合はout.xlsx)'],
                            ["--d", None, self.dir, None, 'ディレクトリ名', 'ディレクトリ指定'],
                            ["--f", None, self.file, None, 'ファイル名', 'ファイル指定'],
                        ])

    

if __name__ == '__main__':
    obj = agrenv()
